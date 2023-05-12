/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const HtmlWebpackTagsPlugin = require("html-webpack-tags-plugin");
const webpack = require("webpack");
const path = require("path");
const dotenv = require("dotenv");
const fs = require("fs");

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
        from: "./src/config.js.template",
        to: "config.js.template",
      },
      {
        from: "./manifest.xml.template",
        to: "manifest.xml.template",
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
    position: "head",
  }),
  new HtmlWebpackTagsPlugin({
    tags: ["launchevent.js"],
    files: ["commands.html"],
    append: true,
    position: "head",
  }),
  new webpack.ProvidePlugin({
    Promise: ["es6-promise", "Promise"],
  }),
];

function replaceEnvPlaceholders(content, outputPath) {
  if (outputPath.endsWith(".xml.template") || outputPath.endsWith(".js.template")) {
    return content
      .toString()
      .replace(/\${BASE_URL}/g, process.env.BASE_URL)
      .replace(/\${WIRE_API_BASE_URL}/g, process.env.WIRE_API_BASE_URL)
      .replace(/\${WIRE_API_AUTHORIZATION_ENDPOINT}/g, process.env.WIRE_API_AUTHORIZATION_ENDPOINT)
      .replace(/\${CLIENT_ID}/g, process.env.CLIENT_ID);
  }

  return content;
}

const pluginsDev = [
  new CopyWebpackPlugin({
    patterns: [
      {
        from: "./manifest.xml.template",
        to: "manifest.xml",
        transform(content, path) {
          return replaceEnvPlaceholders(content, path);
        },
      },
      {
        from: "./src/config.js.template",
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
        from: "./manifest.xml.template",
        to: "manifest.xml",
      },
      {
        from: "./src/config.js.template",
        to: "config.js",
      },
    ],
  }),
];

module.exports = async (env, options) => {
  const isDevelopmentMode = options.mode === "development";

  let envKeys;
  let host, port;
  if (isDevelopmentMode) {
    const envVars = dotenv.config().parsed;
    envKeys = Object.keys(envVars).reduce((prev, next) => {
      prev[`process.env.${next}`] = JSON.stringify(envVars[next]);
      return prev;
    }, {});

    const baseUrl = new URL(process.env.BASE_URL);
    host = baseUrl.hostname;
    port = baseUrl.port;
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
          use: ["ts-loader"],
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
      https:
        fs.existsSync("./devcert/development-key.pem") && fs.existsSync("./devcert/development-cert.pem")
          ? {
              key: fs.readFileSync("./devcert/development-key.pem"),
              cert: fs.readFileSync("./devcert/development-cert.pem"),
            }
          : false,
      host,
      port: port || 8080,
    },
  };

  return config;
};
