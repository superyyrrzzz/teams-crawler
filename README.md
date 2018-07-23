# What is Microsoft Teams Crawler

If you want to export all the chat threads in a Teams channel. Microsoft Graph API can do this [in future releases](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/chatthread). However, you need to hack this for now. This tool do this hack.

# How to use this

## 0. Prerequisite
[Node.js](https://nodejs.org/en/).

## 1. Prepare configuration
create a configuration file named `teams-crawler.json`:
```json
{
  "channelId": "",
  "token": ""
}
```

| Field | Type | Description |
|-------|------|-------------|
| `channelId` | string | The id of a [Teams Channel](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/channel#properties), can get from Microsoft Graph. |
| `token` | string | The [access token](what-is-an-access-token-and-how-do-i-use-it), can get from local Cookies `authtoken`, with value like: `Bearer={TOKEN}&Origin=https://teams.microsoft.com` (only TOKEN part needed)|

## 2. Run and get the results
run:
```
node index.js
```
The result will be saved to `chatThreads.json`, like:
```json
{
  "chatThreads": [
      ...
  ]
}
```

# How this work

This works by mimicing API calls between Teams Online and server. This need to:

1. Log on [Teams Online](https://teams.microsoft.com/)

2. Get the `authtoken` from Cookies. with value like: `Bearer={TOKEN}&Origin=https://teams.microsoft.com`. 

3. Get teams id, **channel id** using teams API
    * List joinedTeams: https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/user_list_joinedteams
    * List channels: https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_list_channels

3. Send request to get first chat thread.
    * URL: https://chatsvcagg.teams.microsoft.com/api/v1/teams/{CHANNEL_ID}/channels/{CHANNEL_ID}?pageSize=20
    * method: GET
    * headers:
        * authorization: `Bearer {TOKEN}`

    This request contains 20 threads as `pageSize` specifies. This is the default age size of Teams Online.

4. Send request to get the remaining threads. This is the same as the above step, with an additional header `x-ms-continuation`. It's value can be found in the response of the previous request keyed `continuationToken`. Resend this repeatedly to get all the threads.