$(document).ready(function () {

    document.getElementById("btnFireup").onclick = function () {
        postPrint();
    };

});

function postPrint() {

    var content = document.getElementById("printContent").innerText;
    var uri = "api/values";
    if (content) {

        $.getJSON(uri)
         .done(function (data) {
             alert('Successed.');
             //// On success, 'data' contains a list of products.
             //$.each(data, function (key, item) {
             //    // Add a list item for the product.
             //    $('<li>', { text: formatItem(item) }).appendTo($('#products'));
             //});
         });

        //$.ajax({
        //    type: "POST",
        //    url: "/api/values",
        //    dataType: "jsonp",
        //    data: { 'content': content},
        //    // beforeSend: function(){},
        //    success: function (result) {
        //        alert('Successed.');
        //    }
        //});
    }
    else {
        alert("Please enter the content to fire up.");
    }

}