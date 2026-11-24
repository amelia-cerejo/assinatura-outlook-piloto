/* Assinatura din�mica - Outlook Exchange Online. Sem bibliotecas externas. */
(function () {
  "use strict";
  const NAME = "Am�lia Cerejo Lopes";
  const PHONE = "966 165 804";
  function esc(v) {return String(v || "").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
  function render(address) {
    const e = esc(address);
    return '<div style="font-family:Arial,sans-serif;font-size:11pt;line-height:1.45;color:#242424">'+
      '<strong>'+esc(NAME)+'</strong><br />'+
      '&#9742; '+esc(PHONE)+'<br />'+
      '&#9993; <a href="mailto:'+e+'" style="color:#1867a4;text-decoration:none">'+e+'</a></div>';
  }
  function finish(event) { try {event.completed();} catch (ignore) {} }
  function update(event, skipWhenNativeSignature) {
    try {
      const item = Office.context.mailbox.item;
      if (!item || !item.from || !item.from.getAsync || !item.body || !item.body.setSignatureAsync) {finish(event); return;}
      const readFrom = () => item.from.getAsync(result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value || !result.value.emailAddress) {finish(event); return;}
        const html = render(result.value.emailAddress);
        item.body.setSignatureAsync(html, {coercionType: Office.CoercionType.Html}, result2 => {
          if (result2.status !== Office.AsyncResultStatus.Succeeded) console.warn('Falha ao definir assinatura', result2.error && result2.error.message);
          finish(event);
        });
      });
      if (skipWhenNativeSignature && typeof item.isClientSignatureEnabledAsync === 'function') {
        item.isClientSignatureEnabledAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value === true) {finish(event); return;}
          readFrom();
        });
      } else {readFrom();}
    } catch (err) {console.error('Assinatura din�mica:', err); finish(event);}
  }
  function onNewMessageComposeHandler(event) { update(event, true); }
  function onMessageFromChangedHandler(event) { update(event, false); }
  Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
  Office.actions.associate('onMessageFromChangedHandler', onMessageFromChangedHandler);
}());