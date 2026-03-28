const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const webpack = require("webpack");
require("dotenv").config();

module.exports = (env, options) => {
  const devMode = options.mode !== "production";

  return {
    devtool: devMode ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
            options: {
              presets: [["@babel/preset-env", { targets: { edge: "18" } }]],
            },
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets", noErrorOnMissing: true }],
      }),
      new webpack.DefinePlugin({
        "process.env.AZURE_CLIENT_ID": JSON.stringify(
          process.env.AZURE_CLIENT_ID || ""
        ),
        "process.env.API_URL": JSON.stringify(
          devMode
            ? "https://localhost:7264/api"
            : "https://ledgerflow-pro.azurewebsites.net/api"
        ),
      }),
    ],
    devServer: {
      https: true,
      port: 3000,
      hot: true,
      headers: { "Access-Control-Allow-Origin": "*" },
      static: { directory: path.join(__dirname, "dist") },
    },
  };
};
