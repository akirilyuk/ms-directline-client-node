/* eslint-disable camelcase */

const axiosRetry = require('axios-retry');
const axios = require('axios').create();
const EventEmitter = require('events');
const WebSocket = require('ws');

// Exponential back-off retry delay between requests
axiosRetry(axios, { retryDelay: axiosRetry.exponentialDelay, retries: 3 });

const EVENTS = {
  ERROR: 'error',
  ACTIVITIES: 'activities',
  WS_CLOSED: 'wsClosed'
};

const ERROR_CODES = {
  wsParsingFailed: 'ws-parsing-failed',
  wsError: 'ws-error',
  wsReconnectErr: 'reconnect-error',
  pollingError: 'polling-error'
};

const defaultConfig = {
  polling: false,
  autoReconnect: true
};

class Conversation extends EventEmitter {
  /**
   * starts a new conversation
   * https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-start-conversation?view=azure-bot-service-3.0
   * @param userId {string} userId of the consumer
   * @param msDirectLineSecret {string} directline secret
   * @param msDirectLineEndpoint {string} directline endpoint
   * @param config {polling:{boolean}, autoReconnect:{boolean}} connector config
   * @returns {Promise<Conversation>} return a new instance, else throws
   */
  static async start({
    userId,
    msDirectLineSecret,
    msDirectLineEndpoint,
    config
  }) {
    const {
      // eslint-disable-next-line camelcase
      data: { conversationId, token, expires_in, streamUrl }
    } = await axios.post(
      `${msDirectLineEndpoint}/conversations`,
      {},
      {
        headers: {
          Authorization: `Bearer ${msDirectLineSecret}`
        }
      }
    );
    return new Conversation({
      conversationId,
      token,
      expires_in,
      userId,
      msDirectLineEndpoint,
      config,
      streamUrl
    });
  }

  /**
   * Create a new Conversation instance
   * @param conversationId {string} ms conversation token
   * @param token {string} auth token
   * @param expires_in {number} expiration time
   * @param streamUrl {string} ws stream url
   * @param msDirectLineEndpoint {string} ms Directline domain
   * @param userId {string} consumer user id
   * @param config {Object} conversation configuration
   */
  constructor({
    conversationId,
    token,
    expires_in,
    streamUrl,
    msDirectLineEndpoint,
    userId,
    config
  }) {
    super();
    this.config = { ...defaultConfig, ...config };
    this.userId = userId;
    this.conversationId = conversationId;
    this.token = token;

    this.expires_in = expires_in;
    this.steamUrl = streamUrl;
    this.msDirectLineEndpoint = msDirectLineEndpoint;
    this.watermarkId = null;
    this.activities = [];
    this.pollingTimer = null;
    this.cleaupState = false;
    if (!this.config.polling) {
      this.createNewWsStream();
    }
  }

  /**
   * remove all external event listeners from this instance
   */
  unbindEvents() {
    Object.keys(EVENTS).forEach(event => {
      this.removeAllListeners(event);
    });
  }

  /**
   * Creates a new WS connection to the provided stream url
   */
  createNewWsStream() {
    this.wsStream = new WebSocket(this.steamUrl);

    this.wsStream.on('closed', async code => {
      if (!this.cleaupState && this.config.autoReconnect) {
        await this.reconnect();
      }
      this.emit(EVENTS.WS_CLOSED, code);
    });

    this.wsStream.on('message', data => {
      if (!data) {
        return;
      }
      try {
        const dataParsed = JSON.parse(data);
        this.processReceivedActivities(dataParsed);
      } catch (error) {
        this.emit(EVENTS.ERROR, { name: ERROR_CODES.wsParsingFailed, error });
      }
    });
    this.wsStream.on('error', async error => {
      this.emit(EVENTS.ERROR, { name: ERROR_CODES.wsError, error });
    });
  }

  /**
   * Reconnect on ws connection close: recreate stream ulr, cleanup old ws, create new ws stream connection
   * @returns {Promise<void>}
   */
  async reconnect() {
    try {
      await this.recreateStreamUrl();
      await this.cleanupWsStream();
      this.createNewWsStream();
    } catch (error) {
      this.emit(EVENTS.ERROR, { name: ERROR_CODES.wsReconnectErr, error });
    }
  }

  /**
   * Cleanup this conversation, stop ws connection, polling, and unbind all event from this instance
   */
  cleanup() {
    this.cleaupState = true;
    if (!this.config.polling && this.wsStream) {
      this.cleanupWsStream();
    }
    this.stopPolling();
    this.unbindEvents();
    this.cleaupState = false;
  }

  /**
   * Cleanup the WS client and remove all listeners
   */
  cleanupWsStream() {
    this.wsStream.terminate();
    this.wsStream.removeAllListeners('open');
    this.wsStream.removeAllListeners('message');
    this.wsStream.removeAllListeners('error');
  }

  /**
   * Recreates a new stream url after ws connection loss and updates this instance with new config values
   * https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-reconnect-to-conversation?view=azure-bot-service-4.0
   * @returns {Promise<void>}
   */
  async recreateStreamUrl() {
    const {
      data: { conversationId, token, streamUrl }
    } = await axios.get(
      `${this.msDirectLineEndpoint}/conversations/${this.conversationId}?watermark=${this.watermarkId}`
    );

    this.conversationId = conversationId;
    this.token = token;
    this.steamUrl = streamUrl;
  }

  /**
   * Send payload to bot, see reference here:
   * https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference?view=azure-bot-service-3.0#activity-object
   *
   * Starts polling if no polling timer exists
   *
   * @param payload {Activity} activity object
   * @param startPolling {boolean} start polling if not started yet, default true
   * @returns {Promise<void>} resolves on success else throws
   */
  async sendPayload(payload, startPolling = true) {
    await axios.post(
      `${this.msDirectLineEndpoint}/conversations/${this.conversationId}/activities`,
      payload,
      {
        headers: {
          Authorization: `Bearer ${this.token}`,
          'Content-Type': 'application/json'
        }
      }
    );

    // start polling after first activity sent from consumer
    if (this.config.polling && !this.pollingTimer && startPolling) {
      await this.startPolling();
    }
  }

  /**
   * Send a text message Activity to the bot
   * @param text {String} contains text to send
   * @returns {Promise<void>} resolves on error else throws
   */
  async sendMessage(text) {
    return this.sendPayload({
      type: 'message',
      from: {
        id: this.userId
      },
      text
    });
  }

  /**
   * Ends the current conversations and cleanups every ongoing data
   * https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-end-conversation?view=azure-bot-service-4.0
   * @param cleanup {boolean} default true, cleanups conversation removes listeners stop polling etc
   * @returns {Promise<void>}
   */
  async endConversation(cleanup = true) {
    const endConvPayload = {
      type: 'endOfConversation',
      from: {
        id: this.userId
      }
    };
    await this.sendPayload(endConvPayload);
    this.activities.push(endConvPayload);
    if (cleanup) {
      this.cleanup();
    }
  }

  /**
   * Poll for next activity responses from MS
   *
   * https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-receive-activities?view=azure-bot-service-3.0#http-get
   *
   * @param direct {boolean} if true, skip new polling timer creation
   * @returns {Promise<null>}
   */
  async pollNextResponse(direct = false) {
    try {
      const { data } = await axios.get(
        `${this.msDirectLineEndpoint}/conversations/${
          this.conversationId
        }/activities${
          this.watermarkId ? `?watermark=${this.watermarkId}` : ''
        }`,
        {
          headers: {
            Authorization: `Bearer ${this.token}`
          }
        }
      );
      await this.processReceivedActivities(data);
      if (!direct) {
        // lazy polling each sec
        this.pollingTimer = setTimeout(this.pollNextResponse.bind(this), 1000);
      }
    } catch (error) {
      // cleanup and stop polling on polling error
      this.emit(EVENTS.ERROR, { name: ERROR_CODES.pollingError, error });
    }

    return null;
  }

  /**
   * Processes received activities data, if polling is enabled and there is fresher data available,
   * will instant trigger new  single out of order poll request
   * @param data
   */
  processReceivedActivities(data) {
    const { activities, watermark } = data;

    // first take care of polling => faster, then do data parsing
    const watermarkId = Number(watermark);

    if (watermarkId > this.watermarkId) {
      // https://docs.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-direct-line-3-0-receive-activities?view=azure-bot-service-3.0#timing-considerations
      // directly poll for next activities and only update data we do not know if case of reconnect
      this.watermarkId = watermarkId;
      if (this.config.polling) {
        this.pollNextResponse(true);
      }
      // add the received activities to the client
      this.activities.push(...activities);
      if (activities.length > 0) {
        this.emit(EVENTS.ACTIVITIES, activities);
      }
    }
  }

  /**
   * Strart polling, promise can be ignored, else resolves after first successful polling else throws
   * @returns {Promise<null>}
   */
  async startPolling() {
    return this.pollNextResponse();
  }

  /**
   * Stops polling for new activities, resume by sending a message or startPolling
   */
  stopPolling() {
    clearTimeout(this.pollingTimer);
    this.pollingTimer = null;
  }

  /**
   * Returns all received activities
   * @returns {[]|Array}
   */
  getActivities() {
    return this.activities;
  }

  /**
   * Get the consumer userId
   * @returns {string}
   */
  getUserId() {
    return this.userId;
  }

  /**
   * Get foreign activities only
   * @returns {Array}
   */
  getForeignActivities() {
    return this.filterOutOwn(this.getActivities());
  }

  /**
   * Filters the passed activities and return only foreign ones
   * @param activities {Array} array with activities
   * @returns {Array}
   */
  filterOutOwn(activities) {
    return activities.filter(activity => activity.from.id !== this.getUserId());
  }
}

Conversation.EVENTS = EVENTS;
Conversation.ERROR_CODES = ERROR_CODES;
module.exports = Conversation;
