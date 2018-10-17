use std::ptr::null_mut;
use winapi::shared::minwindef::DWORD;
use winapi::shared::winerror::{S_FALSE, S_OK};
use winapi::um::combaseapi::{CoInitializeEx, CoInitializeSecurity, CoUninitialize};
use winapi::um::objbase::COINIT_APARTMENTTHREADED;

pub use winapi::shared::rpcdce::{RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_ANONYMOUS};

// uninitialize COM when this drops
pub struct COMInited();

impl COMInited {
    pub fn new(authn_level: DWORD, imp_level: DWORD) -> Result<Self, String> {
        match unsafe { CoInitializeEx(null_mut(), COINIT_APARTMENTTHREADED) } {
            S_OK | S_FALSE => {}
            hr => {
                return Err(format!("CoInitializeEx failed, HRESULT {:#08x}", hr));
            }
        }

        match unsafe {
            CoInitializeSecurity(
                null_mut(), // pSecDesc
                -1,         // cAuthSvc
                null_mut(), // asAuthSvc
                null_mut(), // pReserved1
                authn_level,
                imp_level,
                null_mut(), // pAuthList
                0,          // dwCapabilities
                null_mut(), // pReserved3
            )
        } {
            S_OK => {}
            hr => {
                return Err(format!("CoInitializeSecurity failed, HRESULT {:#08x}", hr));
            }
        }
        Ok(COMInited())
    }
}

impl Drop for COMInited {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

// Handler for return code of HRESULT-returning functions
// TODO: extend to take list of success codes besides S_OK
#[macro_export]
macro_rules! check_hr_ok {
    ($desc:expr, $expr:expr) => {
        match $expr {
            hr @ ::winapi::shared::winerror::S_OK => Ok(hr),
            hr => Err(format!("{} failed, HRESULT = {:#08x}", $desc, hr)),
        }
    };
}

#[macro_export]
macro_rules! check_create {
    ($desc:expr, $cons:expr, $closure:expr) => {{
        let mut obj = ::std::ptr::null_mut();
        let result = check_hr_ok!($desc, $closure(&mut obj));
        result.and_then(|_| $cons($desc, obj))
    }};
}

// TODO: probably should have QueryInterface and CoCreateInstance in here as well
