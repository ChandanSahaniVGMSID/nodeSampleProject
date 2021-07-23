// Copyright (c) Wictor Wilén. All rights reserved. 
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

const gulp = require("gulp");
const package = require("./package.json");
const argv = require("yargs").argv;
const log = require("fancy-log");
const path = require("path");

const config = {};

process.env.VERSION = package.version;

const core = require("yoteams-build-core");

// Initialize core build
core.setup(gulp, config);

process.env['NODE_OPTIONS'] = '--max-old-space-size=4096';
// Add your custom or override tasks below
