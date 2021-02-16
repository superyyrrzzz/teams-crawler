'use strict'

const fs = require('fs');
const https = require('https');

const settings = JSON.parse(fs.readFileSync('teams-crawler.json'));
const teamId = settings.teamId
const channelId = settings.channelId;
const token = settings.token;
const maxRequests = settings.maxRequests != undefined ? settings.maxRequests : -1;

const result = {
  chatThreads: []
};
getChatThreads(maxRequests);

function getOptions(teamId, channelId, token, continuationToken) {
  const options = {
    hostname: 'teams.microsoft.com',
    port: 443,
    path: `/api/csa-msft/api/v1/teams/{teamId}/channels/${channelId}?pageSize=20`,
    headers: {
      authorization: `Bearer ${token}`
    }
  }
  if (continuationToken) {
    options.headers['x-ms-continuation'] = continuationToken;
  }
  return options;
}

function getChatThreads(remaining, continuationToken) {
  if (remaining == 0) {
    console.log(`Max request limits ${maxRequests} reached.`);
    exportResult();
    return;
  }

  let body = '';
  const options = getOptions(teamId, channelId, token, continuationToken)
  const req = https.get(options, (res) => {
    console.log('statusCode:', res.statusCode);

    res.on('data', (d) => {
      body += d;
    });

    res.on('end', (d) => {
      const bodyObj = JSON.parse(body);
      if (bodyObj.replyChains && bodyObj.replyChains.length > 0) {
        result.chatThreads = result.chatThreads.concat(bodyObj.replyChains);
        if (bodyObj.continuationToken) {
          getChatThreads(remaining - 1, bodyObj.continuationToken);
        } else {
          console.log('No continuationToken found in response body.');
          exportResult();
        }
      }
      else {
        console.log('No chat threads found in response body.');
        exportResult();
      }
    })

  })
  req.on('error', (e) => {
    console.error(e);
  })
}

function exportResult() {
  const path = 'chatThreads.json';
  fs.writeFileSync(path, JSON.stringify(result));
  console.log(`Write result to ${path} successfully.`);
}