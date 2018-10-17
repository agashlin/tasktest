use std::ffi::OsString;
use std::fs::File;
use std::io::prelude::*;
use std::ptr::null_mut;
use comical::handle::*;
use winapi::shared::winerror::ERROR_INSUFFICIENT_BUFFER;
use winapi::um::handleapi::INVALID_HANDLE_VALUE;
use winapi::um::minwinbase::LPTR;
use winapi::um::processthreadsapi::{GetCurrentProcess, OpenProcessToken};
use winapi::um::securitybaseapi::{GetTokenInformation, IsValidSid, IsWellKnownSid};
use winapi::um::winbase::LocalAlloc;
use winapi::um::winnt::{
    TokenUser, WinLocalServiceSid, PTOKEN_USER, TOKEN_QUERY, WELL_KNOWN_SID_TYPE,
};

pub fn run(args: &[OsString]) -> Result<(), String> {
    let map = |io| format!("{}", io);

    let mut outfile = File::create("C:\\ProgramData\\tasktest.txt").map_err(map)?;
    outfile.write_fmt(format_args!("OK!\n")).map_err(map)?;
    if is_current_process_user_equal(WinLocalServiceSid)? {
        outfile
            .write_fmt(format_args!("User seems to be LocalService\n"))
            .map_err(map)?;
    } else {
        outfile
            .write_fmt(format_args!("User is not LocalService\n"))
            .map_err(map)?;
    }
    for (i, arg) in args.into_iter().enumerate() {
        match arg.to_str() {
            Some(s) => outfile.write_fmt(format_args!("{}: \"{}\"\n", i, s)),
            None => outfile.write_fmt(format_args!("{}: malformed: {:?}\n", i, arg)),
        }.map_err(map)?;
    }

    Ok(())
}

// Check what user we're running as
fn is_current_process_user_equal(sid_type: WELL_KNOWN_SID_TYPE) -> Result<bool, String> {
    let mut token = INVALID_HANDLE_VALUE;
    check_bool!("OpenProcessToken for current process", unsafe {
        OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &mut token)
    })?;
    let token = HHolder::new("OpenProcessToken", token)?;

    let user_info = {
        let mut needed_length = 0;
        check_bool_expect_err!(
            "GetTokenInformation for size",
            ERROR_INSUFFICIENT_BUFFER,
            unsafe {
                GetTokenInformation(
                    *token,
                    TokenUser,
                    null_mut(),
                    0, // no buffer provided
                    &mut needed_length,
                )
            }
        )?;

        let user_info = LAHolder::new("LocalAlloc for sec token", unsafe {
            LocalAlloc(LPTR, needed_length as usize)
        })?;

        check_bool!("GetTokenInformation for effect", unsafe {
            GetTokenInformation(
                *token,
                TokenUser,
                *user_info,
                needed_length,
                &mut needed_length,
            )
        })?;

        user_info
    };

    check_bool!("IsValidSid", unsafe {
        IsValidSid((*(*user_info as PTOKEN_USER)).User.Sid.clone())
    })?;

    if 0 == unsafe { IsWellKnownSid((*(*user_info as PTOKEN_USER)).User.Sid.clone(), sid_type) } {
        Ok(false)
    } else {
        Ok(true)
    }
}
