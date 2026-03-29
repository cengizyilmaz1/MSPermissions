const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const test = require('node:test');

const {
    buildPermissionMethodsFromDocsTree,
    buildPermissionPowerShellFromDocsTree,
    buildPermissionCodeExamplesFromDocsTree,
    buildResourceDocsFromDocsTree
} = require('../Script/node/lib/graph-docs-parser');

function writeFile(filePath, content) {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    fs.writeFileSync(filePath, content.trimStart(), 'utf8');
}

test('buildPermissionMethodsFromDocsTree parses direct permissions tables', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-direct-'));

    writeFile(path.join(root, 'api-reference', 'v1.0', 'api', 'user-list.md'), `
---
title: "List users"
---

# List users

## Permissions
|Permission type|Permissions (from least to most privileged)|
|:---|:---|
|Delegated (work or school account)|User.ReadBasic.All, User.Read.All|
|Application|User.Read.All, Directory.Read.All|

## HTTP request
\`\`\`http
GET /users
\`\`\`
`);

    const result = buildPermissionMethodsFromDocsTree(root);
    const userReadAll = result.data['User.Read.All'];

    assert.ok(userReadAll);
    assert.equal(userReadAll.v1.length, 1);
    assert.equal(userReadAll.v1[0].path, '/users');
    assert.equal(userReadAll.v1[0].supportsDelegated, true);
    assert.equal(userReadAll.v1[0].supportsApplication, true);
    assert.equal(userReadAll.v1[0].isLeastPrivilege, true);
});

test('buildPermissionMethodsFromDocsTree resolves permission include files', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-include-'));

    writeFile(path.join(root, 'api-reference', 'beta', 'api', 'user-get.md'), `
---
title: "Get user"
---

# Get user

## Permissions
[!INCLUDE [permissions-table](../includes/permissions/user-get-permissions.md)]

## HTTP request
\`\`\`http
GET /me
GET /users/{id | userPrincipalName}
\`\`\`
`);

    writeFile(path.join(root, 'api-reference', 'beta', 'includes', 'permissions', 'user-get-permissions.md'), `
|Permission type|Least privileged permissions|Higher privileged permissions|
|:---|:---|:---|
|Delegated (work or school account)|User.Read|User.Read.All|
|Application|User.Read.All|Directory.Read.All|
`);

    const result = buildPermissionMethodsFromDocsTree(root);
    const userReadAll = result.data['User.Read.All'];

    assert.ok(userReadAll);
    assert.equal(userReadAll.beta.length, 2);
    assert.equal(userReadAll.beta[0].path, '/me');
    assert.equal(userReadAll.beta[1].path, '/users/{id | userPrincipalName}');
    assert.equal(userReadAll.beta[0].supportsDelegated, true);
    assert.equal(userReadAll.beta[0].supportsApplication, true);
    assert.equal(userReadAll.beta[0].docLink, 'https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-beta');
});

test('buildPermissionMethodsFromDocsTree parses single-line permission include tables', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-inline-include-'));

    writeFile(path.join(root, 'api-reference', 'v1.0', 'api', 'chat-list-pinnedmessages.md'), `
---
title: "List pinned messages"
---

# List pinned messages

## Permissions
[!INCLUDE [permissions-table](../includes/permissions/chat-list-pinnedmessages-permissions.md)]

## HTTP request
\`\`\`http
GET /chats/{chat-id}/pinnedMessages
\`\`\`
`);

    writeFile(path.join(root, 'api-reference', 'v1.0', 'includes', 'permissions', 'chat-list-pinnedmessages-permissions.md'), `--- title: "Permissions" description: "Permissions for chat pinned messages." ms.localizationpriority: medium author: "MSGraphDocsVteam" ms.author: "MSGraphDocsVteam" ms.reviewer: "MSGraphDocsVteam" ms.topic: include date: 11/22/2024 --- |Permission type|Least privileged permissions|Higher privileged permissions| |:---|:---|:---| |Application|ChatMessage.Read.All|Not available.|`);

    const result = buildPermissionMethodsFromDocsTree(root);
    const permission = result.data['ChatMessage.Read.All'];

    assert.ok(permission);
    assert.equal(permission.v1[0].path, '/chats/{chat-id}/pinnedMessages');
});

test('buildPermissionPowerShellFromDocsTree parses PowerShell snippets', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-powershell-'));

    writeFile(path.join(root, 'api-reference', 'v1.0', 'api', 'user-get.md'), `
---
title: "Get user"
---

# Get user

## Permissions
|Permission type|Least privileged permissions|Higher privileged permissions|
|:---|:---|:---|
|Application|User.Read.All|Directory.Read.All|

## HTTP request
\`\`\`http
GET /users/{id}
\`\`\`

# [PowerShell](#tab/powershell)
[!INCLUDE [sample-code](../includes/snippets/powershell/get-user-powershell-snippets.md)]
`);

    writeFile(path.join(root, 'api-reference', 'v1.0', 'includes', 'snippets', 'powershell', 'get-user-powershell-snippets.md'), `
\`\`\`powershell
Import-Module Microsoft.Graph.Users
Get-MgUser -UserId $userId
\`\`\`
`);

    const result = buildPermissionPowerShellFromDocsTree(root);
    const userReadAll = result.data['User.Read.All'];

    assert.ok(userReadAll);
    assert.equal(userReadAll.v1.length, 1);
    assert.equal(userReadAll.v1[0].command, 'Get-MgUser');
    assert.equal(userReadAll.v1[0].endpoint, '/users/{id}');
    assert.equal(userReadAll.v1[0].docLink, 'https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0');
});

test('buildPermissionCodeExamplesFromDocsTree parses official language snippets', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-code-'));

    writeFile(path.join(root, 'api-reference', 'v1.0', 'api', 'user-get.md'), `
---
title: "Get user"
---

# Get user

## Permissions
|Permission type|Least privileged permissions|Higher privileged permissions|
|:---|:---|:---|
|Application|User.Read.All|Directory.Read.All|

## HTTP request
\`\`\`http
GET /users/{id}
\`\`\`

# [C#](#tab/csharp)
[!INCLUDE [sample-code](../includes/snippets/csharp/get-user-csharp-snippets.md)]

# [JavaScript](#tab/javascript)
[!INCLUDE [sample-code](../includes/snippets/javascript/get-user-javascript-snippets.md)]

# [Python](#tab/python)
[!INCLUDE [sample-code](../includes/snippets/python/get-user-python-snippets.md)]
`);

    writeFile(path.join(root, 'api-reference', 'v1.0', 'includes', 'snippets', 'csharp', 'get-user-csharp-snippets.md'), `
\`\`\`csharp
var result = await graphClient.Users["{user-id}"].GetAsync();
\`\`\`
`);
    writeFile(path.join(root, 'api-reference', 'v1.0', 'includes', 'snippets', 'javascript', 'get-user-javascript-snippets.md'), `
\`\`\`javascript
const result = await graphClient.users.byUserId('user-id').get();
\`\`\`
`);
    writeFile(path.join(root, 'api-reference', 'v1.0', 'includes', 'snippets', 'python', 'get-user-python-snippets.md'), `
\`\`\`python
result = await graph_client.users.by_user_id('user-id').get()
\`\`\`
`);

    const result = buildPermissionCodeExamplesFromDocsTree(root);
    const userReadAll = result.data['User.Read.All'];

    assert.ok(userReadAll);
    assert.equal(userReadAll.v1.csharp.length, 1);
    assert.equal(userReadAll.v1.javascript.length, 1);
    assert.equal(userReadAll.v1.python.length, 1);
    assert.equal(userReadAll.v1.csharp[0].docLink, 'https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0');
});

test('graph docs parser accepts nested headings and inline code fences', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-inline-'));

    writeFile(path.join(root, 'api-reference', 'v1.0', 'api', 'team-get.md'), `
---
title: "Get team"
---

# Get team

### Permissions for directory readers
|Permission type|Least privileged permissions|Higher privileged permissions|
|:---|:---|:---|
|Application|Team.ReadBasic.All|Team.Read.All|

### HTTP request
\`\`\`http
GET /teams/{team-id}
\`\`\`

### PowerShell
\`\`\`powershell
Import-Module Microsoft.Graph.Teams
Get-MgTeam -TeamId $teamId
\`\`\`

### JavaScript
\`\`\`javascript
const result = await graphClient.teams.byTeamId('team-id').get();
\`\`\`
`);

    const methodsResult = buildPermissionMethodsFromDocsTree(root);
    const powerShellResult = buildPermissionPowerShellFromDocsTree(root);
    const codeExamplesResult = buildPermissionCodeExamplesFromDocsTree(root);

    assert.equal(methodsResult.data['Team.ReadBasic.All'].v1[0].path, '/teams/{team-id}');
    assert.equal(powerShellResult.data['Team.ReadBasic.All'].v1[0].command, 'Get-MgTeam');
    assert.match(codeExamplesResult.data['Team.ReadBasic.All'].v1.javascript[0].code, /graphClient\.teams/);
});

test('buildPermissionMethodsFromDocsTree parses http fences with a space after backticks', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-http-space-'));

    writeFile(path.join(root, 'api-reference', 'v1.0', 'api', 'chat-list-pinnedmessages.md'), `
---
title: "List pinnedChatMessages in a chat"
---

# List pinnedChatMessages in a chat

## Permissions
|Permission type|Least privileged permissions|Higher privileged permissions|
|:---|:---|:---|
|Application|ChatMessage.Read.All|Chat.ReadWrite.All, Chat.Read.All|

## HTTP request
\`\`\` http
GET /chats/{chat-id}/pinnedMessages
\`\`\`
`);

    const result = buildPermissionMethodsFromDocsTree(root);
    const permission = result.data['ChatMessage.Read.All'];

    assert.ok(permission);
    assert.equal(permission.v1[0].path, '/chats/{chat-id}/pinnedMessages');
});

test('buildResourceDocsFromDocsTree parses resource properties, relationships and json representation', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-resource-'));

    writeFile(path.join(root, 'api-reference', 'v1.0', 'resources', 'user.md'), `
---
title: "user resource type"
description: "Represents a Microsoft Entra user account."
---

# user resource type

## Properties
| Property | Type | Description |
|:---|:---|:---|
| id | String | The unique identifier for the user. |
| displayName | String | The name displayed for the user. |

## Relationships
| Relationship | Type | Description |
|:---|:---|:---|
| manager | directoryObject | The user's manager. |

## JSON representation
\`\`\`json
{
  "id": "00000000-0000-0000-0000-000000000000",
  "displayName": "Adele Vance"
}
\`\`\`
`);

    const result = buildResourceDocsFromDocsTree(root);
    const user = result.data.user;

    assert.ok(user);
    assert.equal(user.v1.properties[0].name, 'id');
    assert.equal(user.v1.relationships[0].name, 'manager');
    assert.match(user.v1.jsonRepresentation, /displayName/);
    assert.equal(user.v1.docLink, 'https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0');
});
