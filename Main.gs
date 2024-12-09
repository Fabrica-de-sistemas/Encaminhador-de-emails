/** Número máximo de emails. Após isso, não encaminha nada, somente aviso. */
const CRITICALSPAMNUMBER = 50;

/** Número que emite aviso, mas que depois encaminha e-mail. (este tempo é declarado em MINUTESWAITINGADVISORY */
const ADVISORYSPAMNUMBER = 10; 

/** Se tudo correr bem, esse(s) é(são) o(s) minuto(s) que leva(m) para a próxima verificação */
const MINUTESTRIGGER = 1; 

/** Minutos que vai esperar depois de ter um aviso crítico de segurança */
const MINUTESWAITINGADVISORY = 50;

/** Minutos que vai esperar depois de ter um possível SPAM (definido em CRITICALSPAMNUMBER) */
const MINUTESWAITINGCRITICAL = 120;
const MILITOMINUTE = 1000 * 60; // o tempo é dado em milisegundos

/** O id do arquivo json que está no seu drive  */
const IDFILEJSONEMAILS = "[Insira o id do seu arquivo no drive aqui]";
const EMAILSARRAY = getEmails(IDFILEJSONEMAILS);

/** Domínio que, do arquivo json, quem não tiver, não receberá o email  */
const DOMINIOFORMATTER = "@batatinha.com";

/** Texto que será adicionado previamente no assunto */
const STARTSUBJECTTEXT = "Mensagem reencaminhada do 01-BIT:"

/** O nome do email que aparecerá ao destinatário */
const EMAILNAME = "no-reply 01BIT"

function mailForwarderMain() {
    callMainAfterMinute(MINUTESTRIGGER);  
    try{
      const UNREAD_THREADS_COUNT = GmailApp.getInboxUnreadCount();
      const UNREAD_EMAILS_COUNT = countUnreaded(UNREAD_THREADS_COUNT);

      if (UNREAD_EMAILS_COUNT > CRITICALSPAMNUMBER) {
        let textWarning = `Caixa de e-mail lotada. ${UNREAD_EMAILS_COUNT > UNREAD_THREADS_COUNT? "Verificamos que alguma conversa tem mais de 1 email não lido." : "Há muitos e-mails na caixa."}` 
        sendWarning(textWarning, `${MINUTESWAITINGCRITICAL} minutos`, UNREAD_EMAILS_COUNT); // envia um email de aviso
        callMainAfterMinute(MINUTESWAITINGCRITICAL); // repete o ciclo depois de 2 horas
        
      } else if (UNREAD_EMAILS_COUNT > ADVISORYSPAMNUMBER) {
        
        let textWarning = `Caixa de e-mail congestionada. ${UNREAD_EMAILS_COUNT > UNREAD_THREADS_COUNT? "Verificamos que alguma conversa tem mais de 1 email não lido." : "Há muitos e-mails na caixa."}` 
        sendWarning(textWarning, `${MINUTESWAITINGADVISORY} minutos`, UNREAD_EMAILS_COUNT) // envia um email de aviso
        callMainAfterMinute(MINUTESWAITINGADVISORY, "sendUnreaded");
        
      } else if (UNREAD_EMAILS_COUNT > 0) sendUnreaded() // envia emails não lidos
    }catch(e){
      Logger.log(`Ocorreu o erro abaixo no main. Uma nova tentativa será feita em 20 minutos. \n${e.stack}`);
      callMainAfterMinute(20);
      
    }
}

/**Envia um aviso para todos os emails*/
function sendWarning(warningText, warningTime, unreadCount) {
  let emailsValidos = formatArrayEmails(EMAILSARRAY, DOMINIOFORMATTER).join(",");
  Logger.log(warningText);
  if(!emailsValidos) return null
  const draft = GmailApp.createDraft(
    emailsValidos,
    warningText,
    "",
  {name: EMAILNAME,
  htmlBody: `<h1>Faremos uma pausa no envio de emails</h1>
  <br>
  <p>Como temos ${unreadCount} emails não lidos, faremos uma pausa de ${warningTime} no envio de emails.</p>
  <br>
  ${unreadCount>CRITICALSPAMNUMBER
    ?`<p>Válido ressaltar que esse aviso será disparado novamente se o número de emails não lidos continuar acima de ${CRITICALSPAMNUMBER}. </p>`
    :`<p>Os e-mails serão enviados após esse período de tempo.</p>`}`
  })
  draft.send();


}

/**
* Conta os emails não lidos
*/
function countUnreaded(unreadThreads) {
  let unreadCount = 0
  if (unreadThreads > 0) {
    const threads = GmailApp.getInboxThreads(0, unreadThreads);
    for (const thread of threads) {
      for (const message of thread.getMessages()) {
        if (message.isUnread()) {
          unreadCount++
        }
      }
    }
  }
  return unreadCount
}

function sendUnreaded() {
  const UNREAD_THREADS_COUNT = GmailApp.getInboxUnreadCount();
  const UNREAD_EMAILS_COUNT = countUnreaded(UNREAD_THREADS_COUNT);

  if (UNREAD_EMAILS_COUNT > CRITICALSPAMNUMBER) {
        let textWarning = `Verificamos que foi adicionado mais emails desde o último aviso`
        sendWarning(textWarning, `${MINUTESWAITINGCRITICAL} minutos`, UNREAD_EMAILS_COUNT); // envia um email de aviso
        callMainAfterMinute(MINUTESWAITINGCRITICAL); // repete o ciclo depois de 2 horas
        
  }else{
    Logger.log("Encaminhando e-mails...")
    const unreadCount = GmailApp.getInboxUnreadCount();
    let threads = GmailApp.getInboxThreads(0, unreadCount);

    if (unreadCount > 0) {
      for (const thread of threads) {
        for (const message of thread.getMessages()) {
          if (message.isUnread()) {
            let validatedEmails = formatArrayEmails(EMAILSARRAY, DOMINIOFORMATTER, message.getFrom()).join(",");
            const ACTIVE_EMAIL = Session.getActiveUser().getEmail()
            
            // reencaminha a mensagem
            try{
              if(!message.getFrom().includes(ACTIVE_EMAIL)){
                message.forward(validatedEmails, {subject: `${STARTSUBJECTTEXT} ${message.getSubject()}`, name: "no-reply DA"});  
                message.markRead();
              }
              
              
            }catch(e){
              if(e.message.includes("no recipient")){
                Logger.log(`O email de assunto ${message.getSubject()} não tem pra quem mandar. Marcando o email como lido.`); // caso não tenha para quem mandar
                message.markRead()
              }else Logger.log(`Um erro ao enviar as mensagens não lidas ocorreu:\n ${e}`)
            }
            
          }
          
        }
      }
      threads = GmailApp.getInboxThreads(0, unreadCount);
      for (const thread of threads) {
        for (const message of thread.getMessages()) {
            if(message.getSubject().startsWith(STARTSUBJECTTEXT)) message.moveToTrash();         
        }
      }
    }
    callMainAfterMinute(MINUTESTRIGGER);
  }
}

/**Retorna a lista de arrays de emails válidos. Os parâmetros "emails" e "dominio" recebem provavelmente EMAILSARRAY e DOMINIOFORMATTER (respectivamente). */
function formatArrayEmails(emails, dominio, autor = ""){

  const regex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/; // Expressão regular para encontrar um e-mail
  

  try{
    autor = autor.match(regex)[0]; 
    if (!autor) autor = ""; // Se não tiver email, autor recebe vazio
  }catch {autor = ""};

  let arrayEmails = [];
  for (const email of emails) {
    
    if((email !== autor) && email.includes(dominio)) arrayEmails.push(email);
  }
  return arrayEmails

}



/** cria um novo trigger que se excluirá */
function callMainAfterMinute(minute, functionName = "mailForwarderMain") {
  if(typeof minute !== "number") minute = MINUTESTRIGGER;
    try{
      deleteOldTriggers();
      Logger.log("Criando novo Trigger único");
      ScriptApp
        .newTrigger(functionName)
        .timeBased()
        .after(minute * MILITOMINUTE)
        .create();
      if(ScriptApp.getProjectTriggers().length < 2){
        ScriptApp
          .newTrigger(functionName)
          .timeBased()
          .after((minute + 20) * MILITOMINUTE)
          .create();
        Logger.log(`Atenção! O trigger permanente de redundância não foi detectado. Foi criado então um temporário de redundância.
        Para reverter, basta deletar todos os triggers atuais, criar um trigger permanente de forma manual, executar o arquivo "GetTriggersId" para pegar o ID do trigger permanente e passar na variável "permanentTriggerId`);
      };
    }catch(e){
      Logger.log(`Ocorreu o erro abaixo não tratado ao tentar criar um novo trigger. Uma nova tentativa será feita em 10 minutos. \n${e.message}`);
      deleteOldTriggers();
      callMainAfterMinute(minute * MILITOMINUTE);
  }  
  
}

/** Deleta todos os triggers; é usado principalmente antes de criar um novo trigger. */
function deleteOldTriggers() {
  try{
    for (const trigger of ScriptApp.getProjectTriggers()) {
      if(trigger.getUniqueId() == "1005351431") continue
      ScriptApp.deleteTrigger(trigger);
    }
    Logger.log("Trigger antigos excluídos com sucesso.");
    
  } catch(e){
    Logger.log(e.stack)
    Logger.log(`Ocorreu o erro não tratado: ${e}`)
  }
}

/** retorna um array com a lista de emails */
function getEmails(idfile) {
  try{
    const jFile = DriveApp.getFileById(idfile);
    const jText = jFile.getBlob().getDataAsString();
    let obj = null;

    try {
      obj = JSON.parse(jText);
    } catch (e) {
      Logger.log(e.stack)
      obj = {}
    }

    if (Object.keys(obj).includes("emails")){
      return obj["emails"]
    }
    return [];
  }catch (e) {
    Logger.log(e.stack)
    throw new Error(`Erro ao pegar os emails! Provavelmente arquivo alterado, ou não existe. Verifique o arquivo json ${IDFILEJSONEMAILS}`);
  }
}
