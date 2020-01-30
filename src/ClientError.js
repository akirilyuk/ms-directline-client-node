class ClientError extends Error {
  /**
   * Wrapper for thrown errors, contains the error code as message param
   * @param props
   */
  constructor(props) {
    super(props);
    this.message = props.message;
    this.originalError = props.originalError;
  }
}

module.exports = ClientError;
