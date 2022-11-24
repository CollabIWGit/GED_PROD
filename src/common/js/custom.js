import $ from 'jquery';


$(document).ready(function() {
    if ($(window).width() < 1024) {
        $('.planete-repeated').on('click', function() {
            $('.planete-repeated').removeClass('active');
            $(this).addClass('active');
        });
    }
    else {
        $('.planete-repeated').mouseenter(function() {
            $('.planete-repeated').removeClass('active');
            $(this).addClass('active');
        });
    }
    

    $('.forecast-bloc .forecast').on('click', function() {
        $(this).find('.time').toggleClass('active');
        $(this).find('.degree').toggleClass('active');
    });

    $(".nav-ham").on('click', function(){
        $('.link-header').toggleClass('open');
        $('body').toggleClass('overflow');
        $('.nav-ham .btns').toggleClass('open');
        $('.nav-ham .btns').toggleClass('not-open');
    });

});
  
// var outilSlider = new Swiper('.swiper-container.outils-slider', {
//     slidesPerView: 1,
//     spaceBetween: 30,
//     loop: false,
//     speed: 500,
//     pagination: {
//         el: '.inner-content.outils .swiper-pagination',
//         type: 'bullets',
//         clickable: true
//     },
//     autoplay: {
//         delay: 4500,
//         disableOnInteraction: false,
//     },
// });
  
// var emploiSlider = new Swiper('.swiper-container.emploi-slider', {
//     slidesPerView: 1,
//     spaceBetween: 30,
//     loop: false,
//     speed: 500,
//     pagination: {
//         el: '.info-repeated.emploi .swiper-pagination',
//         type: 'bullets',
//         clickable: true
//     },
//     autoplay: {
//         delay: 4500,
//         disableOnInteraction: false,
//     },
// });
  
// var newSlider = new Swiper('.swiper-container.news-slider', {
//     slidesPerView: 1,
//     spaceBetween: 30,
//     loop: false,
//     speed: 500,
//     pagination: {
//         el: '.info-repeated.news .swiper-pagination',
//         type: 'bullets',
//         clickable: true
//     },
//     autoplay: {
//         delay: 4500,
//         disableOnInteraction: false,
//     },
// });
  
// var familySlider = new Swiper('.swiper-container.family-slider', {
//     slidesPerView: 1,
//     spaceBetween: 30,
//     loop: false,
//     speed: 500,
//     pagination: {
//         el: '.info-repeated.famille .swiper-pagination',
//         type: 'bullets',
//         clickable: true
//     },
//     autoplay: {
//         delay: 4500,
//         disableOnInteraction: false,
//     },
// });