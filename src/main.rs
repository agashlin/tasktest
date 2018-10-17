extern crate winapi;

#[macro_use]
extern crate comical;
extern crate wintask;

mod task;

use std::env;
use std::ffi::OsString;
use std::process;

use comical::bstr::BStr;
use comical::com::{COMInited, RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_ANONYMOUS};

static TASK_NAME: &'static str = "Abalone";

fn main() {
    if let Err(err) = entry() {
        eprintln!("{}", err);
        process::exit(1);
    }
}

fn entry() -> Result<(), String> {
    fn convert_or_fail(os: &OsString) -> Result<&str, String> {
        os.to_str()
            .ok_or_else(|| String::from("Bad command line encoding."))
    }

    let args = env::args_os().collect::<Vec<_>>();

    if args.len() < 2 {
        return Err(String::from("Invaid command line."));
    }

    // TODO: this could probably be a lazy_static
    let com = COMInited::new(RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_ANONYMOUS)?;

    let cmd_args = &args[2..];

    match convert_or_fail(&args[1])? {
        "install" => wintask::install(&com, &BStr::from(TASK_NAME), cmd_args),
        "uninstall" => wintask::uninstall(&com, &BStr::from(TASK_NAME), cmd_args),
        "run" => wintask::on_demand(&com, &BStr::from(TASK_NAME), cmd_args),
        "task" => task::run(cmd_args),
        _ => Err(String::from("Invalid command.")),
    }?;

    Ok(())
}
