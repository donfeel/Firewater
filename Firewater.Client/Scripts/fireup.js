$(document).ready(function () {

    document.getElementById("btnFireup").onclick = function () {
        postPrint();
    };

});

function postPrint() {

    var content = document.getElementById("printContent").innerText;
    if (content) {

        $.ajax({
            type: "POST",
            url: "http://localhost:57316/api/values",
            data: { "": content },
            // beforeSend: function(){},
            success: function (result) {
                alert('Successed.');
            }
        });
    }
    else {
        alert("Please enter the content to fire up.");
    }

}

function tounicode(data) {
    if (data == '') return '请输入汉字';
    var str = '';
    for (var i = 0; i < data.length; i++) {
        str += "\\u" + parseInt(data[i].charCodeAt(0), 10).toString(16);
    }
    return str;
}