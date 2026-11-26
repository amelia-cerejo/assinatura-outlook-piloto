/* Assinatura dinâmica - Outlook Exchange Online. Sem bibliotecas externas. */
(function () {
  "use strict";
  const NAME = "Amélia Cerejo Lopes";
  const PHONE = "966 165 804";
  function esc(v) {return String(v || "").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
  function render(address) {
    const e = esc(address);
    return '<div style="font-family:Arial,sans-serif;font-size:11pt;line-height:1.45;color:#242424">'+
      '<strong>'+esc(NAME)+'</strong><br />'+
      '&#9742; '+esc(PHONE)+'<br />'+
      '&#9993; <a href="mailto:'+e+'" style="color:#1867a4;text-decoration:none">'+e+'</a></div>';
  }
  function finish(event, message) { if (typeof event.report === 'function') event.report(message);
    else event.completed(); }
  function update(event, skipWhenNativeSignature) {
    try {
      const item = Office.context.mailbox.item;
      if (!item || !item.from || !item.from.getAsync || !item.body || !item.body.setSignatureAsync) {finish(event, 'Abre uma mensagem em edição para inserir a assinatura.'); return;}
      const readFrom = () => item.from.getAsync({}, result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value || !result.value.emailAddress) {
          finish(event, (result.error && result.error.message) || 'Não foi possível ler o remetente.');
          return;
        }
        const html = render(result.value.emailAddress);
        item.body.setSignatureAsync(html, {coercionType: Office.CoercionType.Html}, result2 => {
          if (result2.status !== Office.AsyncResultStatus.Succeeded) console.warn('Falha ao definir assinatura', result2.error && result2.error.message);
          finish(event, result2.status === Office.AsyncResultStatus.Succeeded
            ? 'Assinatura inserida para ' + result.value.emailAddress
            : (result2.error && result2.error.message) || 'O Outlook recusou a inserção.');
        });
      });
      if (skipWhenNativeSignature && typeof item.isClientSignatureEnabledAsync === 'function') {
        item.isClientSignatureEnabledAsync({}, result => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value === true) {finish(event); return;}
          readFrom();
        });
      } else {readFrom();}
    } catch (err) {console.error('Assinatura dinâmica:', err); finish(event, err.message);}
  }
  function onNewMessageComposeHandler(event) { update(event, true); }
  function onMessageFromChangedHandler(event) { update(event, false); }
  if (Office.actions && typeof Office.actions.associate === 'function') {
    Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
    Office.actions.associate('onMessageFromChangedHandler', onMessageFromChangedHandler);
  }

  if (typeof document !== 'undefined' && document.getElementById('insert-signature')) {
    const button = document.getElementById('insert-signature');
    const status = document.getElementById('signature-status');
    Office.onReady(() => {
      button.disabled = false;
      button.addEventListener('click', () => {
        status.textContent = 'A inserir assinatura…';
        update({report: message => { status.textContent = message; }}, false);
      });
    });
  }
}());