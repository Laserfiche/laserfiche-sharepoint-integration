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
    '^@laserfiche/types-lf-ui-components$': '<rootDir>/src/__mocks__/@laserfiche/types-lf-ui-components.js',
    '\\.resx$': '<rootDir>/src/__mocks__/genericFileMock.js',
    '^@ms/.*$': '<rootDir>/src/__mocks__/genericFileMock.js',
    '^@microsoft/sp-webpart-base$': '<rootDir>/src/__mocks__/@microsoft/sp-webpart-base.ts',
    '^@microsoft/sp-page-context$': '<rootDir>/src/__mocks__/@microsoft/sp-page-context.ts',
    '^@microsoft/sp-loader$': '<rootDir>/src/__mocks__/@microsoft/sp-loader.ts',
    '^@microsoft/.*$': '<rootDir>/src/__mocks__/genericFileMock.js',
    '\\.png$': '<rootDir>/src/__mocks__/genericFileMock.js',
    '\\.svg$': '<rootDir>/src/__mocks__/genericFileMock.js',
  },
  transform: {
    '^.+\\.(js|jsx|ts|tsx)$': 'babel-jest',
  },
  modulePathIgnorePatterns: ['<rootDir>/lib/'],
  testMatch: ['<rootDir>/src/**/*.(test|spec).(js|jsx|ts|tsx)'],
};