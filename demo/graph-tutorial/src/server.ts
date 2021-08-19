// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <ServerSnippet>
import express from 'express';
import http from 'http';
import fs from 'fs';
import dotenv from 'dotenv';
import path from 'path';

// Load .env file
dotenv.config();

import authRouter from './api/auth';
import graphRouter from './api/graph';
import defaultRouter from './api/default';

const app = express();
const PORT = process.env.PORT||3000;

// Support JSON payloads
app.use(express.json());
app.use(express.static(path.join(__dirname, 'addin')));
app.use(express.static(path.join(__dirname, 'dist/addin')));


app.use('/auth', authRouter);
app.use('/graph', graphRouter);
app.use('/',defaultRouter);
try{
/* const serverOptions = {
  key: fs.readFileSync(process.env.TLS_KEY_PATH!),
  cert: fs.readFileSync(process.env.TLS_CERT_PATH!),
}; *///
//console.log(serverOptions);
/* http.createServer(app).listen(PORT, () => {
  console.log(`⚡️[server]: Server is running at http://localhost:${PORT}`);
}); */

app.listen(PORT, () => {
  console.log(`Example app listening at http://localhost:${PORT}`)
})
}
catch(error){
  console.log(error);
}
// </ServerSnippet>
