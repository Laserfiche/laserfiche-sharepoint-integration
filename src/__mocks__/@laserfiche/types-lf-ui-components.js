// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

// Mock for @laserfiche/types-lf-ui-components used by Jest.
//
// The real package's runtime index.js uses ESM `export var` syntax; Jest's
// default transformIgnorePatterns excludes node_modules from babel-jest, so
// importing the real module throws `SyntaxError: Unexpected token 'export'`.
//
// The package is almost entirely TypeScript declarations (declared classes
// and interfaces), which ts-jest erases at compile time. The only runtime
// values we currently consume in src/ are the LoginState enum values, used
// to compare against `<lf-login>.state`. Values mirror those defined in
// node_modules/@laserfiche/types-lf-ui-components/index.js.
//
// If a future test exercises another runtime export from this package
// (e.g. FieldType, FieldFormat, LoginMode, RedirectBehavior), add it here.

const LoginState = {
  LoggingIn: 'LoggingIn',
  LoggedIn: 'LoggedIn',
  LoggingOut: 'LoggingOut',
  LoggedOut: 'LoggedOut',
};

module.exports = { LoginState };
