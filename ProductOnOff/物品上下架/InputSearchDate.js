function InputSearchDate(ele, fed) {
    var arr = fed.split("_");
    var fd = arr[0];
    var ed = arr[1];
    document.getElementById("giti_cond_updatedtfrom").value = fd;
    document.getElementById("giti_cond_updatedtto").value = ed;
}