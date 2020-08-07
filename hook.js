const express = require('express')
const middleware = require('@line/bot-sdk').middleware
const JSONParseError = require('@line/bot-sdk').JSONParseError
const SignatureValidationFailed = require('@line/bot-sdk').SignatureValidationFailed
const { IncomingWebhook } = require('ms-teams-webhook');
require('dotenv').config()

const app = express()
const url = process.env.MSTEAMS_WEBHOOK;
 
// Initialize
const webhook = new IncomingWebhook(url);

const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESSTOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET
}

app.use(middleware(config))

app.post('/webhook', (req, res) => {
  res.json(req.body.events) // req.body will be webhook event object
  const event = req.body.events[0];
  console.log(event.message.text);
  var messageObj = {
    "@type": "MessageCard",
    "@context": "https://schema.org/extensions",
    "summary": "Issue 176715375",
    "themeColor": "0078D7",
    "title": "Issue opened: \"Chat notifications from LINE\"",
    "sections": [
        {
            "activityTitle": "Sirirat Rungpetcharat",
            "activitySubtitle": "9/13/2016, 11:46am",
            "activityImage": "https://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg",
 
            "text": event.message.text
        }
    ]
  };

  (async () => {
    await webhook.send(JSON.stringify(messageObj)
    );
  })();


if (event.type === 'message') {
  const message = event.message;




  /*if (message.type === 'text' && message.text === 'bye') {
    if (event.source.type === 'room') {
      client.leaveRoom(event.source.roomId);
    } else if (event.source.type === 'group') {
      client.leaveGroup(event.source.groupId);
    } else {
      client.replyMessage(event.replyToken, {
        type: 'text',
        text: 'I cannot leave a 1-on-1 chat!',
      });
    }
  }*/
}
})

app.use((err, req, res, next) => {
  if (err instanceof SignatureValidationFailed) {
    res.status(401).send(err.signature)
    return
  } else if (err instanceof JSONParseError) {
    res.status(400).send(err.raw)
    return
  }
  next(err) // will throw default 500
})

app.listen(8080)