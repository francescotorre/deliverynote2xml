$(document).ready(function(){
    $('#mainNavBar ul li').on('click', function () {
        $('li.active').removeClass('active');
        $(this).addClass('active');
    });
})
