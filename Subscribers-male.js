const addSubscriberButton=$('addOpen');addSubscriberButton.classList.add('toolbar-add');$('extendToggle').insertAdjacentElement('beforebegin',addSubscriberButton);
function masculineText(value){return String(value).replaceAll('اختاري','اختر').replaceAll('حددي','حدد').replaceAll('تريدين','تريد').replaceAll('أضفِ','أضف')}
document.querySelectorAll('#extendPanel span,#extendPanel small').forEach(el=>{el.textContent=masculineText(el.textContent)});
const originalSubscriberToast=toast;toast=message=>originalSubscriberToast(masculineText(message));
const originalSubscriberConfirm=window.confirm.bind(window);window.confirm=message=>originalSubscriberConfirm(masculineText(message));
