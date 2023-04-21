'use strict';

const build = require('@microsoft/sp-build-web');


build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.tslintCmd.enabled = false;
build.initialize(require('gulp'));

build.configureWebpack.mergeConfig({

  additionalConfiguration: (generatedConfiguration) => {

    generatedConfiguration.module.rules.push({

      test: /\.woff2(\?v=[0-9]\.[0-9]\.[0-9])?$/, 

      loader: 'url-loader',

      query: {

        limit: 10000, mimetype: 'application/font-woff2'

      }



    });

    return generatedConfiguration;

  }

});

// build.configureWebpack.mergeConfig({
//   additionalConfig: (generatedConfiguration1) => {
//     generatedConfiguration1.module.rules.push({

//       test: /\.pdf$/,
//       use: [
//         {
//           loader: 'url-loader',
//           options: {
//             limit: 8192, // in bytes
//             name: '[name].[hash].[ext]', // customize output filename
//             outputPath: 'pdfs/', // output path for PDF files
//             publicPath: '/pdfs/', // public URL path for PDF files
//           },
//         },
//       ],
//     });

//     return generatedConfiguration1;
//   }
// });

// module.exports = {
//   module: {
//     rules: [
//       {
//         test: /\.js$/,
//         exclude: /(node_modules|bower_components)/,
//         use: {
//           loader: 'babel-loader',
//           options: {
//             presets: ['@babel/preset-env']
//           }
//         }
//       }
//     ]
//   }
// };



