
function CopyPwd(e,pwd) {
var v = document.getElementById('RandomPassword').value;
document.getElementById('ctl00_ctl00_ctl00_ctl00_ctl05_OldPassword').value = v;
document.getElementById('ctl00_ctl00_ctl00_ctl00_ctl05_NewPassword').value = pwd;
document.getElementById('ctl00_ctl00_ctl00_ctl00_ctl05_ConfirmPassword').value = pwd;
}