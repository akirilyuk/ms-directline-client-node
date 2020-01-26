module.exports = {
  testMatch: ['**/?(*.)+(test).js?(x)'],
  collectCoverage: true,
  collectCoverageFrom: ['src/**/*.js'],
  coveragePathIgnorePatterns: [
    '/node_modules/'
  ],
  testEnvironment: 'node',
  bail: true,
  forceExit: true,
  setupFilesAfterEnv: ['jest-extended']
};
