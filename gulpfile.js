// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

'use strict';

const gulp = require('gulp');
const path = require('path');
const fs = require('fs');

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// Vendor lf-ui-components into lib/Assets/packages/ so SPFx's
// webpack file-loader (rule below) emits content-hashed copies into the
// .sppkg. This avoids the runtime CDN fetch from lfxstatic.com that
// SharePoint's default CSP blocks.
// .css → .cssasset rename keeps SPFx's built-in CSS pipeline from claiming
// the files; the file-loader emits them back as .css.
const PACKAGES_LIB_DIR = path.resolve(__dirname, 'lib/Assets/packages');
const VENDORED_FILES = [
  ['node_modules/@laserfiche/lf-ui-components/cdn/lf-ui-components.js',     'lf-ui-components.js'],
  ['node_modules/@laserfiche/lf-ui-components/cdn/indigo-pink.css',         'indigo-pink.cssasset'],
  ['node_modules/@laserfiche/lf-ui-components/cdn/lf-ms-office-lite.css',   'lf-ms-office-lite.cssasset']
];
build.rig.addPreBuildTask(build.subTask('copy-vendored-packages', function (_g, _o, done) {
  fs.mkdirSync(PACKAGES_LIB_DIR, { recursive: true });
  for (const [src, name] of VENDORED_FILES) {
    fs.copyFileSync(path.resolve(__dirname, src), path.join(PACKAGES_LIB_DIR, name));
  }
  done();
}));

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.module.rules.push(
      {
        test: /\.woff2(\?v=[0-9]\.[0-9]\.[0-9])?$/,
        use: { loader: 'url-loader' }
      },
      {
        test: /\.(js|cssasset)$/,
        include: PACKAGES_LIB_DIR,
        use: {
          loader: 'file-loader',
          options: {
            name: (resourcePath) =>
              '[name].[contenthash:8].' + (resourcePath.endsWith('.cssasset') ? 'css' : 'js'),
            esModule: false
          }
        }
      }
    );
    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
