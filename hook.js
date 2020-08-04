const express = require('express')
const middleware = require('@line/bot-sdk').middleware
const JSONParseError = require('@line/bot-sdk').JSONParseError
const SignatureValidationFailed = require('@line/bot-sdk').SignatureValidationFailed
const { IncomingWebhook } = require('ms-teams-webhook');

const app = express()
const url = 'https://outlook.office.com/webhook/e1c377b2-2568-4fe5-a38e-4673708280b1@dbb514a1-e97b-4b50-be5f-c00508b9ad5a/IncomingWebhook/4d5156c438ff44cfaf1d0b01f8c72e43/700ce299-3bd5-446a-aaa5-43c3ea3e4f8b';
 
// Initialize
const webhook = new IncomingWebhook(url);

const config = {
  channelAccessToken: 'prXb5dd5VhqWQs/SzTxM0pWNbIelTftqz+AiHQP0+Ky+NQ4E1WUIq7bEn++uXpd1eh50nrutaek8X7VNfQDjQTOhS5uwkyzqetbXQpwVW9So0KKg4SZOJzHyqW05jCD+Wim4WaVG9KI2UQXnQ7ws4QdB04t89/1O/w1cDnyilFU=',
  channelSecret: '7c9a6720dca80993cf5fc801f93abe35'
}

app.use(middleware(config))

app.post('/webhook', (req, res) => {
  res.json(req.body.events) // req.body will be webhook event object
  const event = req.body.events[0];

  (async () => {
    await webhook.send(JSON.stringify({
      "@type": "MessageCard",
      "@context": "https://schema.org/extensions",
      "summary": "Issue 176715375",
      "themeColor": "0078D7",
      "title": "Issue opened: \"Push notifications not working\"",
      "sections": [
          {
              "activityTitle": "Miguel Garcie",
              "activitySubtitle": "9/13/2016, 11:46am",
              "activityImage": "https://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg",
   
              "text": "test"
          }
      ]
  })
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