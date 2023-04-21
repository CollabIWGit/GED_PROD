
export const handleTreeResponsive = ()=>{
    
    const togglebtn = document.querySelector(".hamburger");
    togglebtn.addEventListener('click', ()=>{
        document.querySelector(".link-header").classList.toggle("toggled-nav");
    })
        
    const sidebarMenu = document.querySelector("div:has(>#sidebarMenu)");
    const sideToggle = document.querySelector(".left-arrow-responsive");
    const closeSideBar = document.querySelector(".close-sidebar");

    [closeSideBar, sideToggle].forEach(el=>el.addEventListener('click', ()=>sidebarMenu.classList.toggle("sidebar-toggle")));

}