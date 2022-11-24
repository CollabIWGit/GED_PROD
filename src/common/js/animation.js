
import $ from 'jquery';

gsap.registerPlugin(ScrollTrigger);

// !Global fade in
if($('.fadeIn').length > 0) {
    gsap.utils.toArray('.fadeIn').forEach((fadeIn, i) => {
        let tlFadeIn = gsap.timeline({
            scrollTrigger: {
            trigger: fadeIn,
            start: "top 90%",
            }
        });
    tlFadeIn.from(fadeIn.querySelectorAll('.fadeInFirst'), {opacity:0 , duration: 0.5, y:-10})
            .from(fadeIn.querySelectorAll('.fadeInSecond'), {opacity:0 , duration: 0.5, y: -10})
    });
}

if($('.staggerIn').length > 0) {
    gsap.utils.toArray('.staggerIn').forEach((staggerIn, i) => {
        let tlstaggerIn = gsap.timeline({
            scrollTrigger: {
            trigger: staggerIn,
            start: "top center",
            }
        });
    
    tlstaggerIn.from(staggerIn.querySelectorAll('.staggerInFirst'), {opacity:0 , duration: 0.5, y:30, ease: Power1.easeOut})
               .from(staggerIn.querySelectorAll('.staggerInStagger'), {opacity:0 , duration: 0.2, y: 30, stagger: .3, ease: Power2.out},)
      
    });
    
}

if($('.staggerIn2').length > 0) {
    gsap.utils.toArray('.staggerIn2').forEach((staggerIn2, i) => {
        let tlstaggerIn2 = gsap.timeline({
            scrollTrigger: {
            trigger: staggerIn2,
            start: "top 90%",
            // toggleActions: "play reset play reset",
            }
        });
    
    tlstaggerIn2.from(staggerIn2.querySelectorAll('.staggerInFirst2'), {opacity:0 , duration: 0.5, x:-10})
               .from(staggerIn2.querySelectorAll('.staggerInStagger2'), {opacity:0 , duration: 0.5, x: -10,  stagger: .25,},)
      
    });
    
}


/* Smooth scroll */




