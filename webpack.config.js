/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
const { execFileSync } = require("child_process");
const mailNotesConfig = require("./mailnotes.config");

function getBuildTime() {
  const parts = new Intl.DateTimeFormat("de-DE", {
    timeZone: "Europe/Berlin",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).formatToParts(new Date());

  const value = (type) => parts.find((part) => part.type === type)?.value || "";
  return `${value("day")}.${value("month")}.${value("year")} ${value("hour")}:${value("minute")}:${value("second")}`;
}

function getBuildCommit() {
  try {
    const commit = execFileSync("git", ["rev-parse", "--short", "HEAD"], { encoding: "utf8" }).trim();
    const dirty = execFileSync("git", ["status", "--porcelain"], { encoding: "utf8" }).trim() !== "";
    return dirty ? `${commit}-dirty` : commit;
  } catch {
    return "unbekannt";
  }
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return {
    ca: httpsOptions.ca,
    key: httpsOptions.key,
    cert: httpsOptions.cert,
  };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = dev ? mailNotesConfig.development : mailNotesConfig.production;
  const baseUrl = config.baseUrl.endsWith("/") ? config.baseUrl : config.baseUrl + "/";
  const origin = new URL(baseUrl).origin;
  const devOrigin = new URL(mailNotesConfig.development.baseUrl).origin;
  const devUrl = mailNotesConfig.development.baseUrl.endsWith("/")
    ? mailNotesConfig.development.baseUrl
    : mailNotesConfig.development.baseUrl + "/";

  // Im Development-Build spricht das Taskpane nur /api an; webpack übernimmt
  // die Weiterleitung. Im Production-Build wird direkt der lokale Agent benutzt.
  const taskpaneAgentUrl = dev ? "/api" : config.agentUrl;
  const buildTime = getBuildTime();
  const buildCommit = getBuildCommit();

  return {
    devtool: dev ? "source-map" : false,

    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
    },

    output: {
      clean: true,
    },

    resolve: {
      extensions: [".ts", ".html", ".js"],
    },

    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: { loader: "babel-loader" },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
      ],
    },

    plugins: [
      new webpack.DefinePlugin({
        __MAILNOTES_BUILD_MODE__: JSON.stringify(dev ? "development" : "production"),
        __MAILNOTES_AGENT_URL__: JSON.stringify(taskpaneAgentUrl),
        __MAILNOTES_BASE_URL__: JSON.stringify(baseUrl),
        __MAILNOTES_ENABLE_LOGGING__: JSON.stringify(Boolean(config.enableLogging)),
        __MAILNOTES_BUILD_TIME__: JSON.stringify(buildTime),
        __MAILNOTES_BUILD_COMMIT__: JSON.stringify(buildCommit),
      }),

      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),

      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext][query]" },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) return content;

              // Volle URLs auf das Hosting-Verzeichnis umstellen.
              // Den AppDomain-Eintrag dagegen nur auf den Origin setzen.
              return content
                .toString()
                .replace(new RegExp(devUrl.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), baseUrl)
                .replace(new RegExp(devOrigin.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), origin);
            },
          },
        ],
      }),

      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],

    devServer: {
      hot: false,
      liveReload: true,
      headers: { "Access-Control-Allow-Origin": "*" },
      proxy: [
        {
          context: ["/api"],
          target: config.proxyTarget,
          pathRewrite: { "^/api": "" },
          changeOrigin: true,
        },
      ],
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
