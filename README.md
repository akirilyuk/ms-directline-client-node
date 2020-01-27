# ms-directline-client-node
MS Directline v3 Node.js API implementation

[Official Microsoft API reference](https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-concepts?view=azure-bot-service-4.0)


#Introduction

This is a work in progress simple MS Directline V3 API REST/WS Client.

Please note, it just provides a simple wrapper to the MS API, any error will simply be wrapped with an internal error code and forwarded to the error event listener or thrown if possible. FOr more information please see the index.js file.

## Events

```js
const EVENTS = {
  ERROR: 'error', // error event
  ACTIVITIES: 'activities', // receive activities
  WS_CLOSED: 'wsClosed' // handle socket closed
};

// exposed via
Conversation.EVENTS;
```

## Config

```js
const defaultConfig = {
  polling: false, // disable polling per default
  autoReconnect: true // reconnect ws connection on default
};
```
## Error codes

```js
const ERROR_CODES = {
  wsParsingFailed: 'ws-parsing-failed', // parsing of data received via ws failed
  wsError: 'ws-error', // general ws connection error
  wsReconnectErr: 'reconnect-error', // recreating a new ws stream and reconnecting failed
  pollingError: 'polling-error', // failed polling via HTTP GET
  refreshError: 'token-refresh-error', // failed to refresh token
  creationFailed: 'creationFailed' // failed to create object, mostly because conversation call failed
};

// exposed via
Conversation.ERROR_CODES;
```

## Example

```js
require('dotenv').config({ path: `${__dirname}/../.env` });
const Conversation = require('./index');

const delay = ms => {
  return new Promise(resolve => setTimeout(resolve, ms));
};
(async () => {
  try {
    const conversation = await Conversation.start({
      userId: 'test',
      msDirectLineEndpoint: process.env.MS_DIRECTLINE_DOMAIN,
      msDirectLineSecret: process.env.MS_DIRECTLINE_SECRET
    });

    // handle error
    conversation.on('error', console.log);

    await conversation.sendMessage('hi');
    await delay(5000);

    // await response and get activities, usually it takes around 300/500 ms
    // https://docs.microsoft.com/de-de/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-receive-activities?view=azure-bot-service-4.0#timing-considerations
    conversation.getForeignActivities().forEach(activity => {
      if (activity.from.id !== conversation.getUserId()) {
        console.log('found bot response', JSON.stringify(activity));
      }
    });

    // or bind on the activities event and receive them in arrays
    conversation.on(Conversation.EVENTS.ACTIVITIES, activities => {
      // get only foreign ones, note that you can also get them here via the get getForeignActivities method,
      // however it possible that the storage in the conversation already has newer activities stored because of the background polling

      console.log('event activities', activities);
      console.log(
        'conversation cache activities',
        conversation.getActivities()
      );
    });
    await conversation.sendMessage('escalate');
    await delay(2000);
    await conversation.sendMessage('hi');
    await conversation.endConversation(false);
    await delay(1000);
    await conversation.cleanup();
  } catch (err) {
    console.log(err);
  }
})();
```

#Todos

* add full test coverage
* add api documentation in README.MD