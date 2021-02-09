const path = require("path")
const merge = require("webpack-merge")
const ForkTsCheckerWebpackPlugin = require("fork-ts-checker-webpack-plugin")
const SetPublicPathPlugin = require("@rushstack/set-webpack-public-path-plugin").SetPublicPathPlugin
const MiniCSSExtractPlugin = require("mini-css-extract-plugin")

module.exports = merge({
	target: "web",
	entry: {
		"initiative-grid-web-part": path.join(__dirname, "../src/webparts/initiativeGrid/InitiativeGridWebPart.ts"),
		"header-application-customizer": path.join(__dirname, "../src/extensions/header/HeaderApplicationCustomizer.ts"),
		"initiative-form-web-part": path.join(__dirname, "../src/webparts/initiativeForm/InitiativeFormWebPart.ts"),
	},
	output: {
		path: path.join(__dirname, "../dist"),
		filename: "[name].js",
		libraryTarget: "umd",
		library: "[name]",
	},
	performance: {
		hints: false,
	},
	stats: {
		errors: true,
		colors: true,
		chunks: false,
		modules: false,
		assets: false,
	},
	resolve: {
		alias: {
			globalize$: path.resolve(__dirname, "node_modules/globalize/dist/globalize.js"),
			globalize: path.resolve(__dirname, "node_modules/globalize/dist/globalize"),
			cldr$: path.resolve(__dirname, "node_modules/cldrjs/dist/cldr.js"),
			cldr: path.resolve(__dirname, "node_modules/cldrjs/dist/cldr"),
		},
	},
	externals: [
		/^@microsoft\//,
		"InitiativeGridWebPartStrings",
		"HeaderApplicationCustomizerStrings",
		"InitiativeFormWebPartStrings",
	],
	module: {
		rules: [
			{
				test: /\.tsx?$/,
				loader: "ts-loader",
				options: {
					transpileOnly: true,
				},
				exclude: /node_modules/,
			},
			{
				test: /\.(jpg|png|woff|woff2|eot|ttf|svg|gif|dds)$/,
				use: "url-loader?name=[name].[ext]",
			},
			{
				test: /\.css$/,
				use: [
					{
						loader: "@microsoft/loader-load-themed-styles",
						options: {
							async: true,
						},
					},
					{
						loader: "css-loader",
						options: {
							importLoaders: 1,
							modules: true,
						},
					},
					{
						loader: "style-loader",
						options: { singleton: true },
					},
					{
						loader: MiniCSSExtractPlugin.loader,
					},
				],
				include: /\.module\.css$/,
			},
			{
				test: /\.css$/,
				use: ["style-loader", "css-loader"],
				exclude: /\.module\.css$/,
			},
			{
				test: function (fileName) {
					return fileName.endsWith(".module.scss") // scss modules support
				},
				use: [
					{
						loader: "@microsoft/loader-load-themed-styles",
						options: {
							async: true,
						},
					},
					"css-modules-typescript-loader",
					{
						loader: "css-loader",
						options: {
							modules: true,
						},
					}, // translates CSS into CommonJS
					"sass-loader", // compiles Sass to CSS, using Node Sass by default
				],
			},
			{
				test: function (fileName) {
					return !fileName.endsWith(".module.scss") && fileName.endsWith(".scss") // just regular .scss
				},
				use: [
					{
						loader: "@microsoft/loader-load-themed-styles",
						options: {
							async: true,
						},
					},
					{
						loader: "css-loader", // translates CSS into CommonJS
						loader: "sass-loader", // compiles Sass to CSS, using Node Sass by default
					},
				],
			},
		],
	},
	resolve: {
		extensions: [".ts", ".tsx", ".js"],
	},
	plugins: [
		new ForkTsCheckerWebpackPlugin({
			tslint: true,
		}),
		new SetPublicPathPlugin({
			scriptName: {
				name: "[name]_?[a-zA-Z0-9-_]*.js",
				isTokenized: true,
			},
		}),
		new MiniCSSExtractPlugin(),
	],
})
