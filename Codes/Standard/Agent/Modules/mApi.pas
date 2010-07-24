unit mApi;
interface
uses  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,  Dialogs, StdCtrls;
type





// left unchanged ==> public declare function globallock lib 'kernel32' (byval hmem as long) as long;///  memory function
// left unchanged ==> public declare function globalunlock lib 'kernel32' (byval hmem as long) as long;
// left unchanged ==> public declare function globalalloc lib 'kernel32' (byval wflags as long, byval dwbytes as long) as long;
// left unchanged ==> public declare function globalfree lib 'kernel32' (byval hmem as long) as long;
// left unchanged ==> public declare sub globalmemorystatus lib 'kernel32' (lpbuffer as memorystatus);
// left unchanged ==> public declare function getprocessheap lib 'kernel32' () as long;
// left unchanged ==> public declare function heapalloc lib 'kernel32' (byval hheap as long, byval dwflags as long, byval dwbytes as long) as long;
// left unchanged ==> public declare function heapfree lib 'kernel32' (byval hheap as long, byval dwflags as long, lpmem as any) as long;
// left unchanged ==> public declare sub copymemory lib 'kernel32' alias 'rtlmovememory' (hpvdest as any, hpvsource as any, byval cbcopy as long);
// left unchanged ==> public declare sub movememory lib 'kernel32' alias 'rtlmovememory' (destination as any, source as any, byval length as long);
// left unchanged ==> public declare function lstrcpy lib 'kernel32' alias 'lstrcpya' (lpstring1 as any, lpstring2 as any) as long;

// left unchanged ==> public declare function createtoolhelpsnapshot lib 'kernel32' alias 'createtoolhelp32snapshot' (byval lflags as long, lprocessid as long) as long;///  process function
// left unchanged ==> public declare function getcurrentprocess lib 'kernel32' () as long;
// left unchanged ==> public declare function getcurrentprocessid lib 'kernel32' () as long;
// left unchanged ==> public declare function registerserviceprocess lib 'kernel32' (byval dwprocessid as long, byval dwtype as long) as long;
// left unchanged ==> public declare function openprocess lib 'kernel32' (byval dwdesiredaccess as long, byval binherithandle as long, byval dwprocessid as long) as long;
// left unchanged ==> public declare function processfirst lib 'kernel32' alias 'process32first' (byval hsnapshot as long, uprocess as processentry32) as long;
// left unchanged ==> public declare function processnext lib 'kernel32' alias 'process32next' (byval hsnapshot as long, uprocess as processentry32) as long;
// left unchanged ==> public declare function closehandle lib 'kernel32' (byval hobject as long) as long;
// left unchanged ==> public declare function priorityclass+ lib 'kernel32' (byval hprocess as long, byval dwpriorityclass as long);

// left unchanged ==> public declare function computername lib 'kernel32' alias 'computernamea' (byval lpcomputername as string) as long;///  information
// left unchanged ==> public declare function getcomputername lib 'kernel32' alias 'getcomputernamea' (byval lpbuffer as string, nsize as long) as long;
// left unchanged ==> public declare function systemparametersinfo lib 'user32' alias 'systemparametersinfoa' (byval uaction as long, byval uparam as long, byval lpvparam as any, byval fuwinini as long) as long;
// left unchanged ==> public declare function systempowerstate lib 'kernel32' (byval fsuspend as long, byval fforce as long) as long;
// left unchanged ==> public declare function exitwindowsex lib 'user32' (byval uflags as long, byval dwreserved as long) as long;
// left unchanged ==> public declare sub sleep lib 'kernel32' (byval dwmilliseconds as long);
// left unchanged ==> public declare function getversionex lib 'kernel32' alias 'getversionexa' (lpversioninformation as osversioninfo) as long;





// left unchanged ==> public declare function lockwindowupdate lib 'user32' (byval hwndlock as long) as long;/////////////////////////////////////////////////////////////////////////////////////////////////////////
// left unchanged ==> public declare function mapwindowpoints lib 'user32' (byval hwndfrom as long, byval hwndto as long, lppoints as any, byval cpoints as long) as long;
// left unchanged ==> public declare function redrawwindow lib 'user32' (byval hwnd as long, lprcupdate as any, byval hrgnupdate as long, byval flags as long) as boolean;
// left unchanged ==> public declare function iswindowvisible lib 'user32' (byval hwnd as long) as long;
// left unchanged ==> public declare function getclassname lib 'user32' alias 'getclassnamea' (byval hwnd as long, byval lpclassname as string, byval nmaxcount as long) as long;
// left unchanged ==> public declare function getactivewindow lib 'user32' () as long;
// left unchanged ==> public declare function getdesktopwindow lib 'user32' () as long;
// left unchanged ==> public declare function getwindowtext lib 'user32' alias 'getwindowtexta' (byval hwnd as long, byval lpstring as string, byval cch as long) as long;
// left unchanged ==> public declare function getwindow lib 'user32' (byval hwnd as long, byval wcmd as long) as long;
// left unchanged ==> public declare function getwindowrect lib 'user32' (byval hwnd as long, lprect as rect) as long;
// left unchanged ==> public declare function getwindowlong lib 'user32' alias 'getwindowlonga' (byval hwnd as long, byval nindex as long) as long;
// left unchanged ==> public declare function parent lib 'user32' (byval hwndchild as long, byval hwndnewparent as long) as long;
// left unchanged ==> public declare function windowpos lib 'user32' (byval hwnd as long, byval hwndinsertafter as long, byval x as long, byval y as long, byval cx as long, byval cy as long, byval uflags as long) as boolean;
// left unchanged ==> public declare function windowlong lib 'user32' alias 'windowlonga' (byval hwnd as long, byval nindex as long, byval dwnewlong as long) as long;
// left unchanged ==> public declare function findwindow lib 'user32' alias 'findwindowa' (byval lpclassname as string, byval lpwindowname as string) as long;
// left unchanged ==> public declare function findwindowex lib 'user32' alias 'findwindowexa' (byval parenthwnd as long, byval firsthwnd as long, byval lpclassname as string, byval lpwindowname as string) as long;
// left unchanged ==> public declare function showwindow lib 'user32' (byval hwnd as long, byval ncmdshow as long) as long;

// left unchanged ==> public declare function getclientrect lib 'user32' (byval hwnd as long, lprect as any) as boolean;///  geometric function
// left unchanged ==> public declare function getsystemmetrics lib 'user32' (byval nindex as long) as long;

// left unchanged ==> public declare sub shaddtorecentdocs lib 'shell32' (byval uflags as long, byval pv as string);///  shell function
// left unchanged ==> public declare function shemptyrecyclebin lib 'shell32' alias 'shemptyrecyclebina' (byval hwnd as long, byval pszrootpath as string, byval dwflags as long) as long;
// left unchanged ==> public declare function shell_notifyicon lib 'shell32' alias 'shell_notifyicona' (byval dwmessage as long, lpdata as notifyicondataa) as boolean;
// left unchanged ==> public declare function shell_notifyiconw lib 'shell32' (byval dwmessage as long, lpdata as notifyicondataw) as boolean;
// left unchanged ==> public declare function showcursor lib 'user32' (byval bshow as long) as long;
// left unchanged ==> public declare function clipcursor lib 'user32' (lprect as any) as long;
// left unchanged ==> public declare function getcursorpos lib 'user32' (lppoint as pointapi) as long;
// left unchanged ==> public declare function blockinput lib 'user32' (byval fblock as long) as long;

// left unchanged ==> public declare function shappbarmessage lib 'shell32' (byval dwmessage as shappbar_messages, pdata as appbardata) as long;///  system message
// left unchanged ==> public declare function postmessage lib 'user32' alias 'postmessagea' (byval hwnd as long, byval wmsg as long, byval wparam as long, byval lparam as long) as long;
// left unchanged ==> public declare function sendmessage lib 'user32' alias 'sendmessagea' (byval hwnd as long, byval msg as long, byval wparam as long, lparam as any) as long;
// left unchanged ==> public declare function getlasterror lib 'kernel32' () as long;
// left unchanged ==> public declare function formatmessage lib 'kernel32' alias 'formatmessagea' (byval dwflags as long, lpsource as any, byval dwmessageid as long, byval dwlanguageid as long, byval lpbuffer as string, byval nsize as long, arguments as long) as long;





// left unchanged ==> public declare function invalidaterect lib 'user32' (byval hwnd as long, byval lprect as long, byval berase as long) as long;/////////////////////////////////////////////////////////////////////////////////////////////////////////
// left unchanged ==> public declare function selectobject lib 'gdi32' (byval hdc as long, byval hobject as long) as long;
// left unchanged ==> public declare function deleteobject lib 'gdi32' (byval hobject as long) as long;
// left unchanged ==> public declare function getobjectapi lib 'gdi32' alias 'getobjecta' (byval hobject as long, byval ncount as long, lpobject as any) as long;
// left unchanged ==> public declare function getdc lib 'user32' (byval hwnd as long) as long;
// left unchanged ==> public declare function deletedc lib 'gdi32' (byval hdc as long) as long;
// left unchanged ==> public declare function releasedc lib 'user32' (byval hwnd as long, byval hdc as long) as long;
// left unchanged ==> public declare function createcompatibledc lib 'gdi32' (byval hdc as long) as long;
// left unchanged ==> public declare function createcompatiblebitmap lib 'gdi32' (byval hdc as long, byval nwidth as long, byval nheight as long) as long;
// left unchanged ==> public declare function createsolidbrush lib 'gdi32' (byval crcolor as long) as long;
// left unchanged ==> public declare function textcolor lib 'gdi32' (byval hdc as long, byval crcolor as long) as long;
// left unchanged ==> public declare function bkcolor lib 'gdi32' (byval hdc as long, byval crcolor as long) as long;
// left unchanged ==> public declare function bkmode lib 'gdi32' (byval hdc as long, byval nbkmode as long) as long;
// left unchanged ==> public declare function bitblt lib 'gdi32' (byval hdestdc as long, byval x as long, byval y as long, byval nwidth as long, byval nheight as long, byval hsrcdc as long, byval xsrc as long, byval ysrc as long, byval dwrop as long) as long;
// left unchanged ==> public declare function getdevicecaps lib 'gdi32' (byval hdc as long, byval nindex as long) as long;
// left unchanged ==> public const bitspixel = 12;
// left unchanged ==> public const opaque = 2;
// left unchanged ==> public const transparent = 1;





// left unchanged ==> public declare function shfileoperation lib 'shell32' alias 'shfileoperationa' (lpfileop as shfileopstruct) as long;///  file i/o

// left unchanged ==> public declare function enumfonts lib 'gdi32' alias 'enumfontsa' (byval hdc as long, byval lpsz as string, byval lpfontenumproc as long, byval lparam as long) as long;///  common dialogs
// left unchanged ==> public declare function getopenfilename lib 'comdlg32.dll' alias 'getopenfilenamea' (popenfilename as openfilename) as long;

// left unchanged ==> public declare function openprinter lib 'winspool.drv' alias 'openprintera' (byval pprintername as string, phprinter as long, pdefault as any) as long;///  printer
// left unchanged ==> public declare function closeprinter lib 'winspool.drv' (byval hprinter as long) as long;
// left unchanged ==> public declare function enumjobs lib 'winspool.drv' alias 'enumjobsa' (byval hprinter as long, byval firstjob as long, byval nojobs as long, byval level as long, pjob as any, byval cdbuf as long, pcbneeded as long, pcreturned as long) as long;

// left unchanged ==> public declare function netbios lib 'netapi32.dll' (pncb as net_control_block) as byte;///  network

// left unchanged ==> public declare function openprocesstoken lib 'advapi32' (byval processhandle as long, byval desiredaccess as long, tokenhandle as long) as long;///  permission
// left unchanged ==> public declare function lookupprivilegevalue lib 'advapi32' alias 'lookupprivilegevaluea' (byval lpsystemname as string, byval lpname as string, lpluid as luid) as long;
// left unchanged ==> public declare function adjusttokenprivileges lib 'advapi32' (byval tokenhandle as long, byval disableallprivileges as long, newstate as token_privileges, byval bufferlength as long, previousstate as token_privileges, returnlength as long) as long;

// left unchanged ==> public declare function getkeystate lib 'user32' (byval nvirtkey as long) as integer;///  keyboard
// left unchanged ==> public declare function windowshookex lib 'user32' alias 'windowshookexa' (byval idhook as long, byval lpfn as long, byval hmod as long, byval dwthreadid as long) as long;
// left unchanged ==> public declare function callnexthookex lib 'user32' (byval hhook as long, byval ncode as long, byval wparam as long, lparam as any) as long;
// left unchanged ==> public declare function unhookwindowshookex lib 'user32' (byval hhook as long) as long;

// left unchanged ==> public const wm_keydown = $h100;
// left unchanged ==> public const wm_keyup = $h101;
// left unchanged ==> public const wm_syskeydown = $h104;
// left unchanged ==> public const wm_syskeyup = $h105;
// left unchanged ==> public const vk_lcontrol = $ha2;
// left unchanged ==> public const vk_lshift = $ha0;
// left unchanged ==> public const vk_rshift = $ha1;
// left unchanged ==> public const vk_lwin = $h5b;
// left unchanged ==> public const vk_rwin = $h5c;
// left unchanged ==> public const vk_rmenu = $ha5;
// left unchanged ==> public const vk_lmenu = $ha4;
// left unchanged ==> public const vk_tab = $h9;
// left unchanged ==> public const vk_control = $h11;
// left unchanged ==> public const vk_escape = $h1b;
// left unchanged ==> public const hc_action = 0;
// left unchanged ==> public const wh_keyboard_ll = 13;
// left unchanged ==> public const llkhf_altdown = $h20;

// left unchanged ==> public const swp_nomove = 2;
// left unchanged ==> public const swp_nosize = 1;
// left unchanged ==> public const swp_wndflags = swp_nomove or swp_nosize;
// left unchanged ==> public const hwnd_topmost = -1;
// left unchanged ==> public const hwnd_notopmost = -2;

// left unchanged ==> public const hkey_classes_root = $h80000000;
// left unchanged ==> public const hkey_current_config = $h80000005;
// left unchanged ==> public const hkey_current_user = $h80000001;
// left unchanged ==> public const hkey_dyn_data = $h80000006;
// left unchanged ==> public const hkey_local_machine = $h80000002;
// left unchanged ==> public const hkey_performance_data = $h80000004;
// left unchanged ==> public const hkey_users = $h80000003;


// left unchanged ==> public const sw_hide = 0;//// constants for showwindow()
// left unchanged ==> public const sw_normal = 1;
// left unchanged ==> public const sw_showminimized = 2;
// left unchanged ==> public const sw_showmaximized = 3;
// left unchanged ==> public const sw_shownoactivate = 4;
// left unchanged ==> public const sw_show = 5;
// left unchanged ==> public const sw_minimize = 6;
// left unchanged ==> public const sw_showminnoactive = 7;
// left unchanged ==> public const sw_showna = 8;
// left unchanged ==> public const sw_restore = 9;
// left unchanged ==> public const sw_showdefault = 10;
// left unchanged ==> public const normal_priority_class = $h20;
// left unchanged ==> public const idle_priority_class = $h40;
// left unchanged ==> public const high_priority_class = $h80;
// left unchanged ==> public const realtime_priority_class = $h100;
// left unchanged ==> public const process_dup_handle = $h40;
// left unchanged ==> public const process_all_access = 0;
// left unchanged ==> public const th32cs_snapprocess as long = 2+;
// left unchanged ==> public const max_path+ = 260;

// left unchanged ==> public const format_message_allocate_buffer = $h100;
// left unchanged ==> public const format_message_from_system = $h1000;
// left unchanged ==> public const lang_neutral = $h0;
// left unchanged ==> public const sublang_default = $h1;

// left unchanged ==> public const ewx_logoff = 0;
// left unchanged ==> public const ewx_shutdown = 1;
// left unchanged ==> public const ewx_reboot = 2;
// left unchanged ==> public const ewx_force = 4;

// left unchanged ==> public const nim_add = $h0;
// left unchanged ==> public const nim_modify = $h1;
// left unchanged ==> public const nim_delete = $h2;

// left unchanged ==> public const nif_message = $h1;
// left unchanged ==> public const nif_icon = $h2;
// left unchanged ==> public const nif_tip = $h4;

// left unchanged ==> public const rdw_invalidate = $h1;
// left unchanged ==> public const rdw_allchildren = $h80;
// left unchanged ==> public const rdw_updatenow = $h100;

// left unchanged ==> public const swp_nozorder = $h4;
// left unchanged ==> public const swp_framechanged = $h20        '  the frame changed: send wm_nccalcsize;
// left unchanged ==> public const swp_drawframe = swp_framechanged;
// left unchanged ==> public const swp_hidewindow = $h80;
// left unchanged ==> public const swp_showwindow = $h40;

// left unchanged ==> public const sherb_noconfirmation = $h1;
// left unchanged ==> public const sherb_noprogressui = $h2;
// left unchanged ==> public const sherb_nosound = $h4;

// left unchanged ==> public const wm_command = $h111;
// left unchanged ==> public const wm_mousemove = $h200;
// left unchanged ==> public const wm_lbuttondown = $h201;
// left unchanged ==> public const wm_lbuttonup = $h202;
// left unchanged ==> public const wm_rbuttondblclk = $h206;
// left unchanged ==> public const wm_rbuttondown = $h204;
// left unchanged ==> public const wm_rbuttonup = $h205;
// left unchanged ==> public const wm_redraw = $hb;
// left unchanged ==> public const wm_user as long = $h400;
// left unchanged ==> public const wm_myhook as long = wm_user + 1;

// left unchanged ==> public const min_all = 419;
// left unchanged ==> public const min_all_undo = 416;

// left unchanged ==> public const gwl_style = (-16);
// left unchanged ==> public const gwl_exstyle = (-20);

// left unchanged ==> public const gw_hwndfirst = 0;
// left unchanged ==> public const gw_hwndlast = 1;
// left unchanged ==> public const gw_hwndnext = 2;
// left unchanged ==> public const gw_hwndprev = 3;
// left unchanged ==> public const gw_max = 5;
// left unchanged ==> public const gw_owner = 4;

// left unchanged ==> public const ws_border = $h800000;
// left unchanged ==> public const ws_ex_staticedge = $h20000;

// left unchanged ==> public const spi_screensaverrunning = 97;
// left unchanged ==> public const spi_deskwallpaper = 20;
// left unchanged ==> public const spif_sendwininichange = $h2;
// left unchanged ==> public const spif_updateinifile = $h1;

// left unchanged ==> public const sm_cycaption = 4;
// left unchanged ==> public const sm_cxscreen = 0;
// left unchanged ==> public const sm_cyscreen = 1;

// left unchanged ==> public const fo_copy = $h2;
// left unchanged ==> public const fo_delete = $h3;
// left unchanged ==> public const fo_move = $h1;
// left unchanged ==> public const fo_rename = $h4;

// left unchanged ==> public const fof_allowundo = $h40;
// left unchanged ==> public const fof_confirmmouse = $h2;
// left unchanged ==> public const fof_filesonly = $h80;
// left unchanged ==> public const fof_multidestfiles = $h1;
// left unchanged ==> public const fof_no_connected_elements = $h2000;
// left unchanged ==> public const fof_noconfirmation = $h10;
// left unchanged ==> public const fof_noconfirmmkdir = $h200;
// left unchanged ==> public const fof_nocopysecurityattribs = $h800;
// left unchanged ==> public const fof_noerrorui = $h400;
// left unchanged ==> public const fof_norecursion = $h1000;
// left unchanged ==> public const fof_renameoncollision = $h8;
// left unchanged ==> public const fof_silent = $h4;
// left unchanged ==> public const fof_simpleprogress = $h100;
// left unchanged ==> public const fof_wantmappinghandle = $h20;
// left unchanged ==> public const fof_wantnukewarning = $h4000;

// left unchanged ==> public const cf_printerfonts = $h2;
// left unchanged ==> public const cf_screenfonts = $h1;
// left unchanged ==> public const cf_both = (cf_screenfonts or cf_printerfonts);
// left unchanged ==> public const cf_effects = $h100&;
// left unchanged ==> public const cf_forcefontexist = $h10000;
// left unchanged ==> public const cf_inittologfontstruct = $h40&;
// left unchanged ==> public const cf_limitsize = $h2000&;
// left unchanged ==> public const regular_fonttype = $h400;

// left unchanged ==> public const fw_normal = 400;
// left unchanged ==> public const default_char = 1;
// left unchanged ==> public const out_default_precis = 0;
// left unchanged ==> public const clip_default_precis = 0;
// left unchanged ==> public const default_quality = 0;
// left unchanged ==> public const default_pitch = 0;
// left unchanged ==> public const ff_roman = 16;
// left unchanged ==> public const lf_facesize = 32;

// left unchanged ==> public const gmem_moveable = $h2;
// left unchanged ==> public const gmem_zeroinit = $h40;
// left unchanged ==> public const heap_zero_memory as long = $h8;
// left unchanged ==> public const heap_generate_exceptions as long = $h4;

// left unchanged ==> public const cchdevicename = 32;
// left unchanged ==> public const cchformname = 32;

// left unchanged ==> public const job_status_paused = $h1;
// left unchanged ==> public const job_status_error = $h2;
// left unchanged ==> public const job_status_deleting = $h4;
// left unchanged ==> public const job_status_spooling = $h8;
// left unchanged ==> public const job_status_printing = $h10;
// left unchanged ==> public const job_status_offline = $h20;
// left unchanged ==> public const job_status_paperout = $h40;
// left unchanged ==> public const job_status_printed = $h80;
// left unchanged ==> public const job_status_deleted = $h100;
// left unchanged ==> public const job_status_blocked_devq = $h200;
// left unchanged ==> public const job_status_user_intervention = $h400     ' windows 95 only;

// left unchanged ==> public const no_priority = 0;
// left unchanged ==> public const max_priority = 99;
// left unchanged ==> public const min_priority = 1;
// left unchanged ==> public const def_priority = 1;

// left unchanged ==> public const ncbastat as long = $h33;
// left unchanged ==> public const ncbnamsz as long = 16;
// left unchanged ==> public const ncbre as long = $h32;


type kbdllhookstruct = record
vkcode : longint;
scancode : longint;
flags : longint;
time : longint;
dwextrainfo : longint;
end;

type osversioninfo = record
dwosversioninfosize : longint;
dwmajorversion : longint;
dwminorversion : longint;
dwbuildnumber : longint;
dwplatformid : longint;
szcsdversion : string;// maintenance string for pss usage
end;

type luid = record
usedpart : longint;
ignoredfornowhigh32bitpart : longint;
end;

type token_privileges = record
privilegecount : longint;
theluid : luid;
attributes : longint;
end;

type processentry32 = record
dwsize : longint;
cntusage : longint;
th32processid : longint;
th32defaultheapid : longint;
th32moduleid : longint;
cntthreads : longint;
th32parentprocessid : longint;
pcpriclassbase : longint;
dwflags : longint;
szexefile : string;
end;

type memorystatus
dwlength : longint;
dwmemoryload : longint;
dwtotalphys : longint;
dwavailphys : longint;
dwtotalpagefile : longint;
dwavailpagefile : longint;
dwtotalvirtual : longint;
dwavailvirtual : longint;
end;

type notifyicondataa
cbsize : longint;
hwnd : longint;
uid : longint;
uflags : longint;
ucallbackmessage : longint;
hicon : longint;
sztip : string;
end;

type notifyicondataw
cbsize : longint;
hwnd : longint;
uid : longint;
uflags : longint;
ucallbackmessage : longint;
hicon : longint;
sztip(0 to 127) as byte
end;

// left unchanged ==> public enum shappbar_messages;
abm_new := $h0;
abm_remove := $h1;
abm_querypos := $h2;
abm_pos := $h3;
abm_getstate := $h4;
abm_gettaskbarpos := $h5;
abm_activate := $h6;
abm_getautohidebar := $h7;
abm_autohidebar := $h8;
abm_windowposchanged := $h9;
end;

// left unchanged ==> public enum shappbar_notifications;
abn_statechange := $h0;
abn_poschanged := $h1;
abn_fullscreenapp := $h2;
abn_windowarrange := $h3;
end;

// left unchanged ==> public enum shappbar_states;
abs_autohide := $h1;
abs_alwaysontop := $h2;
end;

// left unchanged ==> public enum shappbar_edges;
abe_left := 0;
abe_top := 1;
abe_right := 2;
abe_bottom := 3;
end;

type rect
left : longint;
top : longint;
right : longint;
bottom : longint;
end;

type pointapi
x : longint;
y : longint;
end;

type appbardata
cbsize : longint;
hwnd : longint;
ucallbackmessage : longint;
uedge : shappbar_edges;
rc : rect;
lparam : longint;
end;

type shfileopstruct
hwnd : longint;
wfunc : longint;
pfrom : string;
pto : string;
fflags : integer;
faborted : boolean;
hnamemaps : longint;
sprogress : string;
end;

type logfont = record
lfheight : longint;
lfwidth : longint;
lfescapement : longint;
lforientation : longint;
lfweight : longint;
lfitalic : byte;
lfunderline : byte;
lfstrikeout : byte;
lfchar : byte;
lfoutprecision : byte;
lfclipprecision : byte;
lfquality : byte;
lfpitchandfamily : byte;
lffacename: array[1..lf_facesize] of byte;
end;


type openfilename = record
lstructsize : longint;
hwndowner : longint;
hinstance : longint;
lpstrfilter : string;
lpstrcustomfilter : string;
nmaxcustfilter : longint;
nfilterindex : longint;
lpstrfile : string;
nmaxfile : longint;
lpstrfiletitle : string;
nmaxfiletitle : longint;
lpstrinitialdir : string;
lpstrtitle : string;
flags : longint;
nfileoff : integer;
nfileextension : integer;
lpstrdefext : string;
lcustdata : longint;
lpfnhook : longint;
lptemplatename : string;
end;

type systemtime = record
wyear : integer;
wmonth : integer;
wdayofweek : integer;
wday : integer;
whour : integer;
wminute : integer;
wsecond : integer;
wmilliseconds : integer;
end;


type net_control_block = record//ncb
ncb_command    as byte
ncb_retcode    as byte
ncb_lsn        as byte
ncb_num        as byte
ncb_buffer     as long
ncb_length     as integer
ncb_callname   as string * ncbnamsz
ncb_name       as string * ncbnamsz
ncb_rto        as byte
ncb_sto        as byte
ncb_post       as long
ncb_lana_num   as byte
ncb_cmd_cplt   as byte
ncb_reserve: array[1..9] of byte;
ncb_event      as long// reserved, must be 0
end;

type adapter_status = record
adapter_address: array[1..5] of byte;
rev_major         as byte
reserved0         as byte
adapter_type      as byte
rev_minor         as byte
duration          as integer
frmr_recv         as integer
frmr_xmit         as integer
iframe_recv_err   as integer
xmit_aborts       as integer
xmit_success      as long
recv_success      as long
iframe_xmit_err   as integer
recv_buff_unavail : integer;
t1_timeouts       as integer
ti_timeouts       as integer
reserved1         as long
free_ncbs         as integer
max_cfg_ncbs      as integer
max_ncbs          as integer
xmit_buf_unavail  as integer
max_dgram_size    as integer
pending_sess      as integer
max_cfg_sess      as integer
max_sess          as integer
max_sess_pkt_size : integer;
name_count        as integer
end;

type name_buffer = record
name        as string * ncbnamsz
name_num    as integer
name_flags  as integer
end;

type astat = record
adapt          as adapter_status
namebuff(30)   as name_buffer
end;

end.

