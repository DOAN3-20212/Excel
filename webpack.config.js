/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

/* global require, module, process, __dirname */

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { cacert: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      functions: "./src/functions/functions.js",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
      // popup2: "./src/dialogs/popup2.js"
    },
    output: {
      devtoolModuleFilenameTemplate: "webpack:///[resource-path]?[loaders]",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      fallback: {
        cardinal: false,
        buffer: false,
        crypto: false,
        net: false,
        url: false,
        stream: false,
        timers: false,
        tls: false,
        zlib: false,
      }
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-env"],
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.js",
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "popup.html",
        template: "./src/dialogs/popup.html",
        chunks: ["polyfill", "popup"]
      }),
      new HtmlWebpackPlugin({
        filename: "ProductPlant.html",
        template: "./src/dialogs/ProductPlant.html",
        chunks: ["polyfill", "popup"]
      }),
      new HtmlWebpackPlugin({
        filename: "Worker.html",
        template: "./src/dialogs/Worker.html",
        chunks: ["polyfill", "popup"]
      }),
      new HtmlWebpackPlugin({
        filename: "Date.html",
        template: "./src/dialogs/Date.html",
        chunks: ["polyfill", "popup"]
      }),
      new HtmlWebpackPlugin({
        filename: "task.html",
        template: "./src/dialogs/task.html",
        chunks: ["polyfill", "popup"]
      })
    ],
    devServer: {
      static: [__dirname],
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
