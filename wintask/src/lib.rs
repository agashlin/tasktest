#[macro_use]
extern crate winapi;
#[macro_use]
extern crate comical;

pub mod taskschd;

use taskschd::*;

use std::ffi::OsString;
use std::os::windows::ffi::OsStrExt;
use std::ops::Deref;
use std::ptr::null_mut;
use comical::bstr::*;
use comical::handle::*;
use comical::com::*;
use comical::variant::*;
use winapi::ctypes::c_void;
use winapi::shared::minwindef::{DWORD, LPVOID, MAX_PATH};
use winapi::shared::wtypesbase::CLSCTX_INPROC_SERVER;
use winapi::um::combaseapi::CoCreateInstance;
use winapi::um::processthreadsapi::GetCurrentProcess;
use winapi::um::winbase::QueryFullProcessImageNameW;
use winapi::Interface;

define_unsafe_com_holder!(TaskService, ITaskService);
define_unsafe_com_holder!(TaskFolder, ITaskFolder);
define_unsafe_com_holder!(TaskDefinition, ITaskDefinition);
define_unsafe_com_holder!(TaskSettings, ITaskSettings);
define_unsafe_com_holder!(RegistrationInfo, IRegistrationInfo);
define_unsafe_com_holder!(RegisteredTask, IRegisteredTask);
define_unsafe_com_holder!(ActionCollection, IActionCollection);
define_unsafe_com_holder!(Action, IAction);
define_unsafe_com_holder!(ExecAction, IExecAction);
define_unsafe_com_holder!(RunningTask, IRunningTask);
define_unsafe_com_holder!(IdleSettings, IIdleSettings);

fn connect_task_service() -> Result<(TaskService, TaskFolder), String> {
    /*
    let task_service = {
        let mut task_service = null_mut();
        check_hr_ok!("CoCreateInstance of TaskScheduler", unsafe {
        TaskService::new("CoCreateInstance of TaskScheduler", task_service)?
    };
    */

    let task_service = check_create!(
        "CoCreateInstance of TaskScheduler",
        TaskService::new,
        |task_service| unsafe {
            CoCreateInstance(
                &CLSID_TaskScheduler,
                null_mut(), // pUnkOuter
                CLSCTX_INPROC_SERVER,
                &ITaskService::uuidof(),
                task_service as *mut *mut ITaskService as *mut LPVOID,
            )
        }
    )?;

    check_hr_ok!("ITaskService::Connect", unsafe {
        let null = Variant::null().get();
        (**task_service).Connect(
            null, // serverName
            null, // user
            null, // domain
            null, // password
        )
    })?;

    let root_folder = check_create!("Get root folder", TaskFolder::new, |root_folder| unsafe {
        (**task_service).GetFolder(BStr::from("\\").get(), root_folder)
    })?;

    Ok((task_service, root_folder))
}

// TODO This should probably be parameterized more.
pub fn install(_: &COMInited, task_name: &BStr, args: &[OsString]) -> Result<(), String> {
    if args.len() != 0 {
        return Err(String::from("Expected no args for install"));
    }

    let mut current_proc_name = [0u16; MAX_PATH + 1];
    // NOTE: this shadows the old current_proc_name with a reference to a slice of it
    let current_proc_name = {
        let mut current_proc_name_len = MAX_PATH as DWORD;

        check_bool!("QueryFullProcessImageNameW", unsafe {
            QueryFullProcessImageNameW(
                GetCurrentProcess(),
                0, // dwFlags
                current_proc_name.as_mut_ptr(),
                &mut current_proc_name_len as *mut DWORD,
            )
        })?;

        ingest_ws(&current_proc_name, current_proc_name_len)?
    };

    let task_def;
    let root_folder;
    {
        let (task_service, rf) = connect_task_service()?;
        root_folder = rf;

        // If the same task exists, remove it. Allowed to fail.
        unsafe { (**root_folder).DeleteTask(task_name.get(), 0) };

        task_def = check_create!("Create new task", TaskDefinition::new, |task_def| unsafe {
            (**task_service).NewTask(
                0, // flags (reserved)
                task_def,
            )
        })?;
    }

    {
        let reg_info = check_create!(
            "get_RegistrationInfo",
            RegistrationInfo::new,
            |info| unsafe { (**task_def).get_RegistrationInfo(info) }
        )?;

        check_hr_ok!("put_Author", unsafe {
            (**reg_info).put_Author(BStr::from("Mozilla").get())
        })?;
    }

    {
        let settings = check_create!("get_Settings", TaskSettings::new, |s| unsafe {
            (**task_def).get_Settings(s)
        })?;

        check_hr_ok!("put_MultipleInstances", unsafe {
            (**settings).put_MultipleInstances(TASK_INSTANCES_IGNORE_NEW)
        })?;

        check_hr_ok!("put_AllowDemandStart", unsafe {
            (**settings).put_AllowDemandStart(VARIANT_TRUE)
        })?;

        check_hr_ok!("put_RunOnlyIfIdle", unsafe {
            (**settings).put_RunOnlyIfIdle(VARIANT_FALSE)
        })?;

        check_hr_ok!("put_DisallowStartIfOnBatteries", unsafe {
            (**settings).put_DisallowStartIfOnBatteries(VARIANT_FALSE)
        })?;

        check_hr_ok!("put_StopIfGoingOnBatteries", unsafe {
            (**settings).put_StopIfGoingOnBatteries(VARIANT_FALSE)
        })?;

        let idle_settings = check_create!("get_IdleSettings", IdleSettings::new, |s| unsafe {
            (**settings).get_IdleSettings(s)
        })?;

        check_hr_ok!("put_StopOnIdleEnd", unsafe {
            (**idle_settings).put_StopOnIdleEnd(VARIANT_FALSE)
        })?;
    }

    {
        let action_collection = check_create!("get_Actions", ActionCollection::new, |ac| unsafe {
            (**task_def).get_Actions(ac)
        })?;

        let action = check_create!("Create Action", Action::new, |a| unsafe {
            (**action_collection).Create(TASK_ACTION_EXEC, a)
        })?;

        let exec_action =
            check_create!("QueryInterface IExecAction", ExecAction::new, |ea| unsafe {
                (**action).QueryInterface(
                    &IExecAction::uuidof(),
                    ea as *mut *mut IExecAction as *mut *mut c_void,
                )
            })?;

        check_hr_ok!("Set exec action path", unsafe {
            (**exec_action).put_Path(bstr_from_u16(current_proc_name.as_slice()).get())
        })?;

        check_hr_ok!("Set exec action args", unsafe {
            (**exec_action).put_Arguments(BStr::from("task $(Arg0)").get())
        })?;
    }

    let _v;
    {
        let br = &mut BStr::from("");
        _v = unsafe { Variant::wrap(br).get() };
    }

    let registered_task =
        check_create!("RegisterTaskDefinition", RegisteredTask::new, |rt| unsafe {
            (**root_folder).RegisterTaskDefinition(
                task_name.get(),
                *task_def,
                TASK_CREATE_OR_UPDATE as i32,
                Variant::wrap(&mut BStr::from("NT AUTHORITY\\LocalService")).get(),
                Variant::null().get(),
                TASK_LOGON_SERVICE_ACCOUNT,
                Variant::wrap(&mut BStr::empty()).get(), // sddl
                rt,
            )
        })?;

    // Allow read and execute access by builtin users, this is required to Get the task and
    // call Run on it
    check_hr_ok!("SetSecurityDescriptor", unsafe {
        (**registered_task).SetSecurityDescriptor(
            BStr::from("D:(A;;GRGX;;;BU)").get(),
            TASK_DONT_ADD_PRINCIPAL_ACE as i32,
        )
    })?;

    Ok(())
}

pub fn uninstall(_: &COMInited, task_name: &BStr, args: &[OsString]) -> Result<(), String> {
    if args.len() != 0 {
        return Err(String::from("Expected no args for uninstall"));
    }

    let (_, root_folder) = connect_task_service()?;
    check_hr_ok!("DeleteTask", unsafe {
        (**root_folder).DeleteTask(task_name.get(), 0)
    })?;

    Ok(())
}

pub fn on_demand(_com: &COMInited, task_name: &BStr, args: &[OsString]) -> Result<(), String> {
    let text;
    if args.len() == 1 {
        text = args[0].encode_wide().collect::<Vec<u16>>();
    } else {
        text = vec![0];
    }

    let (_, root_folder) = connect_task_service()?;

    let task = check_create!("GetTask", RegisteredTask::new, |t| unsafe {
        (**root_folder).GetTask(task_name.get(), t)
    })?;

    let _running_task = check_create!("Run Task", RunningTask::new, |rt| unsafe {
        (**task).Run(Variant::wrap(&bstr_from_u16(&text)).get(), rt)
    })?;

    Ok(())
}
