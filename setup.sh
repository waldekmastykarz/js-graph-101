#!/usr/bin/env bash

# login
npx -p @pnp/cli-microsoft365@next -- m365 login --authType browser

# create AAD app
npx -p @pnp/cli-microsoft365@next -- m365 aad app add --name "Graph JS 101" --multitenant --redirectUris "https://localhost" --platform spa

# write app to env.js
