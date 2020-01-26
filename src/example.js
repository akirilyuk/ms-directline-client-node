require('dotenv').config();
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
    await conversation.endConversation(true);
  } catch (err) {
    console.log(err);
  }
})();
