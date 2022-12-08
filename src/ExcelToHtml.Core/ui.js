(function () {
    var containers = document.querySelectorAll(".e2h-main .e2h-container");
    var current = 0;
    var tabs = document.querySelectorAll(".e2h-main .e2h-tabs .e2h-tab-item");
    tabs.forEach((el, idx) => {
        el.addEventListener("click", evt => {
            console.log(idx, evt);
            containers[current].classList.remove("e2h-active");
            //containers[current].classList.add("e2h-hide");
            //containers[idx].classList.remove("e2h-hide");
            containers[idx].classList.add("e2h-active");

            tabs[current].classList.remove("e2h-tab-active");
            el.classList.add("e2h-tab-active");
            current = idx;


        });
    });
}());