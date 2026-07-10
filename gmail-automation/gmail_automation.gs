// ===== 設定 =====
const FOLDER_ID = '1CMs5ScgWWH6okg698e3b2_oaqI0Wx0l3';
const REPLY_THRESHOLD_HOURS = 24;

/**
 * 請求書メールの添付ファイルをGoogle Driveに保存する
 */
function saveInvoicesToDrive() {
  const folder = DriveApp.getFolderById(FOLDER_ID);

  const query = '(subject:請求書 OR subject:invoice) newer_than:30d';
  const threads = GmailApp.search(query);

  const processedIds = getProcessedMessageIds_();
  let savedCount = 0;

  for (const thread of threads) {
    for (const message of thread.getMessages()) {
      const msgId = message.getId();
      if (processedIds.has(msgId)) continue;

      const attachments = message.getAttachments();
      for (const attachment of attachments) {
        const name = attachment.getName();
        if (!/\.(pdf|xlsx?|csv|png|jpe?g)$/i.test(name)) continue;

        const dateStr = Utilities.formatDate(message.getDate(), 'Asia/Tokyo', 'yyyyMMdd');
        const safeName = `${dateStr}_${name}`;

        const existing = folder.getFilesByName(safeName);
        if (!existing.hasNext()) {
          folder.createFile(attachment).setName(safeName);
          savedCount++;
          console.log(`保存完了: ${safeName}`);
        }
      }

      markMessageProcessed_(msgId);
    }
  }

  console.log(`請求書保存: ${savedCount}件`);
}

/**
 * 未返信メール・添付ファイル付きメールをまとめて通知する
 */
function checkAndNotify() {
  const myEmail = Session.getEffectiveUser().getEmail();
  const unanswered = getUnansweredEmails_(myEmail);
  const withAttachments = getEmailsWithAttachments_();

  if (unanswered.length === 0 && withAttachments.length === 0) {
    console.log('通知対象なし');
    return;
  }

  sendDigestNotification_(unanswered, withAttachments, myEmail);
  console.log(`通知送信: 未返信${unanswered.length}件、添付付き${withAttachments.length}件`);
}

/**
 * 24時間以上返信していないメール一覧を返す
 */
function getUnansweredEmails_(myEmail) {
  const now = new Date();
  const cutoff = new Date(now.getTime() - REPLY_THRESHOLD_HOURS * 3600 * 1000);

  const query = 'in:inbox -category:promotions -category:social -category:updates newer_than:7d';
  const threads = GmailApp.search(query);

  const unanswered = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    const lastMsg = messages[messages.length - 1];

    if (lastMsg.getFrom().toLowerCase().includes(myEmail.toLowerCase())) continue;
    if (isAutomatedEmail_(lastMsg)) continue;
    if (lastMsg.getDate() >= cutoff) continue;

    unanswered.push({
      subject: thread.getFirstMessageSubject(),
      from: lastMsg.getFrom(),
      received: lastMsg.getDate(),
      url: `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}`
    });
  }

  return unanswered;
}

/**
 * 添付ファイル付きメール一覧を返す
 */
function getEmailsWithAttachments_() {
  const query = 'in:inbox has:attachment -category:promotions -category:social -category:updates newer_than:7d';
  const threads = GmailApp.search(query);

  const result = [];

  for (const thread of threads) {
    const messages = thread.getMessages();
    const lastMsg = messages[messages.length - 1];

    if (isAutomatedEmail_(lastMsg)) continue;

    const attachmentNames = [];
    for (const message of messages) {
      for (const att of message.getAttachments()) {
        attachmentNames.push(att.getName());
      }
    }

    if (attachmentNames.length > 0) {
      result.push({
        subject: thread.getFirstMessageSubject(),
        from: lastMsg.getFrom(),
        received: lastMsg.getDate(),
        attachments: attachmentNames,
        url: `https://mail.google.com/mail/u/0/#inbox/${thread.getId()}`
      });
    }
  }

  return result;
}

/**
 * まとめ通知メールを送信する
 */
function sendDigestNotification_(unanswered, withAttachments, toEmail) {
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  const sections = [];
  if (unanswered.length > 0) sections.push(`未返信${unanswered.length}件`);
  if (withAttachments.length > 0) sections.push(`添付付き${withAttachments.length}件`);

  const mailSubject = `【要確認】${sections.join('・')} (${now})`;

  let body = '';

  if (unanswered.length > 0) {
    body += `■ ${REPLY_THRESHOLD_HOURS}時間以上返信していないメール（${unanswered.length}件）\n`;
    body += '─'.repeat(50) + '\n\n';

    unanswered.forEach((item, i) => {
      body += `【${i + 1}】${item.subject}\n`;
      body += `　送信者 : ${item.from}\n`;
      body += `　受信日時: ${Utilities.formatDate(item.received, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}\n`;
      body += `　リンク  : ${item.url}\n\n`;
    });
  }

  if (withAttachments.length > 0) {
    if (body) body += '\n';
    body += `■ 添付ファイル付きメール（${withAttachments.length}件）\n`;
    body += '─'.repeat(50) + '\n\n';

    withAttachments.forEach((item, i) => {
      body += `【${i + 1}】${item.subject}\n`;
      body += `　送信者 : ${item.from}\n`;
      body += `　受信日時: ${Utilities.formatDate(item.received, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}\n`;
      body += `　添付   : ${item.attachments.join('、')}\n`;
      body += `　リンク  : ${item.url}\n\n`;
    });
  }

  GmailApp.sendEmail(toEmail, mailSubject, body);
}

/**
 * メルマガ・自動通知メールかどうかを判定する
 */
function isAutomatedEmail_(message) {
  const from = message.getFrom().toLowerCase();
  const subject = message.getSubject().toLowerCase();

  const patterns = [
    /no.?reply/,
    /noreply/,
    /donotreply/,
    /newsletter/,
    /notification/,
    /メルマガ/,
    /配信停止/,
    /unsubscribe/,
    /auto.?reply/,
    /自動返信/,
    /bounce/,
    /mailer-daemon/,
    /info@/,
    /news@/,
    /support@/,
    /system@/,
  ];

  return patterns.some(p => p.test(from) || p.test(subject));
}

// ===== 処理済みメッセージIDの管理（重複保存防止） =====

function getProcessedMessageIds_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('processedMsgIds') || '[]';
  return new Set(JSON.parse(raw));
}

function markMessageProcessed_(id) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('processedMsgIds') || '[]';
  const ids = JSON.parse(raw);
  if (!ids.includes(id)) {
    ids.push(id);
    props.setProperty('processedMsgIds', JSON.stringify(ids.slice(-2000)));
  }
}

/**
 * トリガーに設定する関数（両方まとめて実行）
 * 0:00 / 6:00 / 12:00 / 18:00 のみ実行
 */
function runAll() {
  const hour = parseInt(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'H'));
  if (![0, 12].includes(hour)) return;
  saveInvoicesToDrive();
  checkAndNotify();
}
