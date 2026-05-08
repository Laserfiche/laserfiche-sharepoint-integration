// Copyright (c) Laserfiche.
// Licensed under the MIT License. See LICENSE.md in the project root for license information.

'use strict';

const gulp = require('gulp');
const path = require('path');

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// Pre-build task: copy the lf-ui-components CDN assets and zone.js into a local
// folder. The folder is gitignored. Files are emitted into the .sppkg via the
// file-loader rule below so they load from SharePoint's own CDN at runtime,
// satisfying Microsoft's default Content Security Policy (no external scripts).
//
// CSS files are copied with a `.cssasset` extension so SPFx's built-in CSS
// loader pipeline does not pick them up. Our file-loader rule emits them back
// with the original `.css` extension via the `name` option.
const fs = require('fs');
const copyLfUiComponents = build.subTask('copy-lf-ui-components', function (localGulp, buildOptions, done) {
  const destDir = path.resolve(__dirname, 'src/Assets/lf-ui-components');
  if (!fs.existsSync(destDir)) {
    fs.mkdirSync(destDir, { recursive: true });
  }
  const copies = [
    ['node_modules/@laserfiche/lf-ui-components/cdn/lf-ui-components.js', 'lf-ui-components.js'],
    ['node_modules/@laserfiche/lf-ui-components/cdn/indigo-pink.css', 'indigo-pink.cssasset'],
    ['node_modules/@laserfiche/lf-ui-components/cdn/lf-ms-office-lite.css', 'lf-ms-office-lite.cssasset'],
    ['node_modules/zone.js/bundles/zone.umd.min.js', 'zone.umd.min.js']
  ];
  for (const [src, destName] of copies) {
    fs.copyFileSync(path.resolve(__dirname, src), path.join(destDir, destName));
  }
  done();
});
build.rig.addPreBuildTask(copyLfUiComponents);

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    const lfUiAssetsLibPath = path.resolve(__dirname, 'lib/Assets/lf-ui-components');

    generatedConfiguration.module.rules.push(
      {
        test: /\.woff2(\?v=[0-9]\.[0-9]\.[0-9])?$/,
        use: {
          loader: 'url-loader'
        }
      },
      {
        test: /\.(js|cssasset)$/,
        include: lfUiAssetsLibPath,
        use: {
          loader: 'file-loader',
          options: {
            name: (resourcePath) => {
              const ext = resourcePath.endsWith('.cssasset') ? 'css' : 'js';
              return '[name].[contenthash:8].' + ext;
            },
            esModule: false
          }
        }
      }
    );
    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
