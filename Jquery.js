$(document).ready(function () {
    //документ готов

    //если нажат элемент класса .subButton, включить экран загрузки
    $(".subButton").click(function () {
        $(".loadingDiv").css("visibility", "visible");        
    });
});

