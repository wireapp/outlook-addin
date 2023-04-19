/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const HtmlWebpackTagsPlugin = require("html-webpack-tags-plugin");
const webpack = require("webpack");
const path = require("path");
const dotenv = require("dotenv");

const plugins = [
  new CopyWebpackPlugin({
    patterns: [
      {
        // for event-based-activation
        from: "./src/launchevent/launchevent.js",
        to: "launchevent.js",
      },
      {
        from: "assets/*",
        to: "assets/[name][ext][query]",
      },
      {
        from: "./src/config.template.js",
        to: "config.template.js",
      },
      {
        from: "./manifest.template.xml",
        to: "manifest.template.xml",
      },
    ],
  }),
  new HtmlWebpackPlugin({
    filename: "taskpane.html",
    template: "./src/taskpane/taskpane.html",
    chunks: ["taskpane", "vendor", "polyfills"],
    scriptLoading: "blocking",
    inject: "head",
  }),
  new HtmlWebpackPlugin({
    filename: "commands.html",
    template: "./src/commands/commands.html",
    chunks: ["commands"],
    scriptLoading: "blocking",
    inject: "head",
  }),
  new HtmlWebpackPlugin({
    filename: "authorize.html",
    template: "./src/authorize/authorize.html",
    chunks: ["authorize"],
    scriptLoading: "blocking",
    inject: "head",
  }),
  new HtmlWebpackPlugin({
    filename: "callback.html",
    template: "./src/callback/callback.html",
    chunks: ["callback"],
    scriptLoading: "blocking",
    inject: "head",
  }),
  new HtmlWebpackTagsPlugin({
    tags: ["config.js"],
    append: false,
    publicPath: false,
    position: "head",
  }),
  new webpack.ProvidePlugin({
    Promise: ["es6-promise", "Promise"],
  }),
];

function replaceEnvPlaceholders(content, outputPath) {
  if (outputPath.endsWith(".xml") || outputPath.endsWith(".js")) {
    return content
      .toString()
      .replace(/\${ADDIN_HOST}/g, process.env.ADDIN_HOST + (process.env.ADDIN_PORT ? ":" + process.env.ADDIN_PORT : ""))
      .replace(/\${API_HOST}/g, process.env.API_HOST)
      .replace(/\${AUTHORIZE_HOST}/g, process.env.AUTHORIZE_HOST)
      .replace(/\${CLIENT_ID}/g, process.env.CLIENT_ID);
  }

  return content;
}

const pluginsDev = [
  new CopyWebpackPlugin({
    patterns: [
      {
        from: "./manifest.template.xml",
        to: "manifest.xml",
        transform(content, path) {
          return replaceEnvPlaceholders(content, path);
        },
      },
      {
        from: "./src/config.template.js",
        to: "config.js",
        transform(content, path) {
          return replaceEnvPlaceholders(content, path);
        },
      },
    ],
  }),
];

const pluginsDocker = [
  new CopyWebpackPlugin({
    patterns: [
      {
        from: "./manifest.template.xml",
        to: "manifest.xml",
      },
      {
        from: "./src/config.template.js",
        to: "config.js",
      },
    ],
  }),
];

module.exports = async (env, options) => {
  const isDevelopmentMode = options.mode === "development";

  let envKeys;
  if (isDevelopmentMode) {
    const envVars = dotenv.config().parsed;
    envKeys = Object.keys(envVars).reduce((prev, next) => {
      prev[`process.env.${next}`] = JSON.stringify(envVars[next]);
      return prev;
    }, {});
  }

  const config = {
    mode: options.mode,
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      vendor: ["react", "react-dom", "core-js", "@fluentui/react"],
      taskpane: ["react-hot-loader/patch", "./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
      authorize: "./src/authorize/authorize.ts",
      callback: "./src/callback/callback.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: ["@babel/preset-typescript"],
            },
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["react-hot-loader/webpack", "ts-loader"],
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
      ...(isDevelopmentMode ? [new webpack.DefinePlugin(envKeys)] : []),
      ...(isDevelopmentMode ? pluginsDev : pluginsDocker),
      ...plugins,
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: options.https !== undefined ? options.https : env.WEBPACK_BUILD,
      },
      host: process.env.ADDIN_HOST,
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
