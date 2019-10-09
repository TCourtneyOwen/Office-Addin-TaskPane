// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in.
*/
const https = require('https')
const express = require('express');
const bodyParser =require('body-parser');
const cors = require('cors');
const devCerts = require('office-addin-dev-certs');
const env = process.env.NODE_ENV || 'development';


/* Create the express app and add the required middleware */
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
app.use(express.static('dist'))
/* Turn off caching when debugging */
app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    next();
});

startServer(app, env);

async function startServer(app, env) {
    if (env === 'development') {
        const options = await devCerts.getHttpsServerOptions();
        https.createServer(options, app).listen(3000, () => console.log('Server running on 3000'));
    }
    else {
        app.listen(process.env.port || 1337, () => console.log(`Server listening on port ${process.env.port}`));
    }
}
