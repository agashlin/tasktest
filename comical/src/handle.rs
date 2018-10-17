use std::error::Error;
use std::ops::Deref;
use std::ptr::null_mut;
use widestring;
use winapi::um::handleapi::{CloseHandle, INVALID_HANDLE_VALUE};
use winapi::um::winbase::LocalFree;
use winapi::um::winnt::{HANDLE, PVOID};

// A simple container with Drop and Deref.
// TODO Probably would be better as a trait?
#[macro_export]
macro_rules! define_unsafe_holder {
    ($holder_type:ident, $held_type:ty, $drop_fn:expr, $valid:expr, $empty:expr) => {
        pub struct $holder_type($held_type);

        impl $holder_type {
            pub fn new(desc: &'static str, v: $held_type) -> Result<$holder_type, String> {
                if $valid(&v) {
                    Ok($holder_type(v))
                } else {
                    Err(String::from(desc))
                }
            }

            pub fn valid(&self) -> bool {
                $valid(&self.0)
            }
        }

        impl Drop for $holder_type {
            fn drop(&mut self) {
                if self.valid() {
                    unsafe {
                        $drop_fn(self.0);
                        self.0 = $empty;
                    };
                }
            }
        }

        impl Deref for $holder_type {
            type Target = $held_type;

            fn deref(&self) -> &$held_type {
                &self.0
            }
        }
    };
}

define_unsafe_holder!(
    LAHolder,
    PVOID,
    LocalFree,
    |p: &PVOID| !p.is_null(),
    null_mut()
);
define_unsafe_holder!(
    HHolder,
    HANDLE,
    CloseHandle,
    |p: &HANDLE| *p != INVALID_HANDLE_VALUE,
    INVALID_HANDLE_VALUE
);
/*
define_unsafe_holder!(
    CoTaskMemHolder,
    winnt::PVOID,
    combaseapi::CoTaskMemFree,
    |p: &winnt::PVOID| !p.is_null(),
    null_mut()
);
*/

// This probably shouldn't be a macro, but I'm a uncomfortable about casting IWhatever to
// IUnknown in order to use Release. TODO find a better way?
#[macro_export]
macro_rules! define_unsafe_com_holder {
    ($holder_type:ident, $held_type:ty) => {
        define_unsafe_holder!(
            $holder_type,
            *mut $held_type,
            |p: *mut $held_type| (*p).Release(),
            |p: &*mut $held_type| !(*p as *mut $held_type).is_null(),
            null_mut()
        );
    };
}

// Handler for functions returning zero on failure
// TODO Probably better as a fn?
#[macro_export]
macro_rules! check_bool {
    ($desc:expr, $expr:expr) => {
        match $expr {
            0 => {
                let last_error = unsafe { ::winapi::um::errhandlingapi::GetLastError() };
                Err(format!("{} failed, last error = {}", $desc, last_error))
            }
            _ => Ok(()),
        }
    };
}

#[macro_export]
macro_rules! check_bool_expect_err {
    ($desc:expr, $err:expr, $expr:expr) => {
        match $expr {
            0 => {
                let last_error = unsafe { ::winapi::um::errhandlingapi::GetLastError() };
                if last_error == $err {
                    Ok(())
                } else {
                    Err(format!("{} failed, last error = {}", $desc, last_error))
                }
            }
            _ => Err(format!("{} didn't fail", $desc)),
        }
    };
}

pub fn ingest_ws<'a>(s: &'a [u16], len: u32) -> Result<&'a widestring::WideCStr, String> {
    if len as usize >= usize::max_value() {
        return Err(String::from("Length out of range"));
    }
    let len = len as usize + 1; // include terminator
    if len > s.len() {
        return Err(String::from("Wide string longer than slice"));
    }
    match widestring::WideCStr::from_slice_with_nul(&s[0..len]) {
        Err(e) => Err(String::from(e.description())),
        Ok(cstr) => Ok(cstr),
    }
}
