// filepath: jest.config.js
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jest-environment-jsdom',
  setupFilesAfterEnv: ['./setupTests.ts'],
  moduleNameMapper: {
    '\\.(css|less|scss|sass)$': 'identity-obj-proxy',
    '^@laserfiche/lf-repository-api-client$': '<rootDir>/src/__mocks__/@laserfiche/lf-repository-api-client.js',
    '^@laserfiche/lf-js-utils$': '<rootDir>/src/__mocks__/@laserfiche/lf-js-utils.js',
    '^@laserfiche/lf-ui-components-services$': '<rootDir>/src/__mocks__/@laserfiche/lf-ui-components-services.js',
    '\\.resx$': '<rootDir>/src/__mocks__/genericFileMock.js',
    '^@ms/.*$': '<rootDir>/src/__mocks__/genericFileMock.js',
    '^@microsoft/sp-webpart-base$': '<rootDir>/src/__mocks__/@microsoft/sp-webpart-base.ts',
    '^@microsoft/.*$': '<rootDir>/src/__mocks__/genericFileMock.js',
    '\\.png$': '<rootDir>/src/__mocks__/genericFileMock.js',

  },
  transform: {
    '^.+\\.(js|jsx|ts|tsx)$': 'babel-jest', // Add this line to transform all JavaScript and TypeScript files
  },
  testMatch: ['<rootDir>/src/**/*.(test|spec).(js|jsx|ts|tsx)'], // Add this line to specify the test match pattern
};