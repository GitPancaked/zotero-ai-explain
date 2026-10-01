# Secure API-key storage smoke test

Use disposable, restricted API keys. Do not publish key values, credential output, or profile
backups. Install the v0.4.2 XPI in a disposable Zotero profile, keeping production profiles
untouched.

## Windows

1. Open settings. Confirm the storage message names Windows Credential Manager.
2. Save an OpenAI, Anthropic, or Gemini key. Confirm Save completes and API requests work.
3. Restart Zotero. Confirm the provider can authenticate without reentering the key.
4. Inspect Windows Credential Manager: the entry uses `zotero-ai-explain/<provider>`.
5. Confirm no raw key appears in the corresponding Zotero API-key preference or `prefs.js` after
   shutdown.
6. Clear the key field and Save. Confirm the OS credential is removed and remains absent after
   restart.

## Arch Linux

1. Install `libsecret` plus a Secret Service backend, e.g. GNOME Keyring. Start the service through
   your desktop session and unlock a password-protected default collection. Do not use an empty
   password.
2. Repeat the save/restart/delete checks above. Inspect entry attributes in your keyring manager:
   `service=zotero-ai-explain`, `provider=openai|anthropic|gemini`.
3. Stop the service, remove `secret-tool` from the test environment, or cancel an unlock prompt.
   Confirm saving a new key displays an error and keeps the form open; no key should appear in
   prefs.
4. Restore/unlock the keyring, retry Save, and confirm success.

## Migration and failure recovery

1. With Zotero closed, prepare a disposable profile with a legacy API-key preference.
2. Start Zotero with available secure storage. Confirm the key migrates, authentication works, and
   the legacy pref is removed after shutdown.
3. Repeat with unavailable storage. Confirm the legacy pref is preserved, a migration warning is
   shown, and direct API authentication does not silently use the plaintext fallback.
4. Unlock/restore storage and restart. Confirm migration succeeds without losing the key.
5. If a secure credential already exists, confirm it wins over a stale legacy pref.

Keys are currently shared by this plugin across all Zotero profiles of the same OS user. Use one
profile at a time for these tests. Remove disposable credentials after testing. Migration does not
clean old backups, and OS stores do not protect keys from malware running as the same user.
