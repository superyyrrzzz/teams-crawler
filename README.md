# Teams Crawler

## What is Teams Crawler

This tool crawls all the chat threads from a teams channel by calling the API that Teams online uses.

## How to use this

### 0. Prerequisite

[Node.js](https://nodejs.org/en/).

### 1. Prepare configuration

Create a configuration file named `teams-crawler.json`:

```json
{
  "teamId": "",
  "channelId": "",
  "token": ""
}
```

| Field | Type | Description |
|-------|------|-------------|
| `teamId` | string | The if of a Team, can get from Microsoft Graph. |
| `channelId` | string | The id of a [Teams Channel](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/channel#properties), can get from Microsoft Graph. |
| `token` | string | The access token, can get by inspect the traffic in Teams Online and get from the request header. (I do not find another easier way to get the token. I tried get an access token from Graph API, but it cannot work here...)

### 2. Run and get the results

run:

```ps
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

## How this work

This works by mimicing API calls between Teams Online and server. This need to:

1. Log on [Teams Online](https://teams.microsoft.com/)

2. Get the `authtoken` from Cookies. with value like: `Bearer={TOKEN}&Origin=https://teams.microsoft.com`. 

3. Get teams id, **channel id** using teams API

    * List joinedTeams: <https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/user_list_joinedteams>

    * List channels: <https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_list_channels>

4. Send request to get first chat thread.
    * URL: </api/csa-msft/api/v1/teams/{teamId}/channels/${channelId}?pageSize=20>
    * method: GET
    * headers:
        * authorization: `Bearer {TOKEN}`

    This request contains 20 threads as `pageSize` specifies. This is the default age size of Teams Online.

5. Send request to get the remaining threads. This is the same as the above step, with an additional header `x-ms-continuation`. It's value can be found in the response of the previous request keyed `continuationToken`. Resend this repeatedly to get all the threads.