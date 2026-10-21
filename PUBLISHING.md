# GitHub Actions Trusted Publishing for @otomasi/office

This repository holds ONLY the already-public Windows Office agent npm source, NOT private MCP server, Office OAuth database, secrets, or per-user credentials.

GitHub workflow: `.github/workflows/publish.yml`
npm package: [@otomasi/office](https://www.npmjs.com/package/@otomasi/office)

## One-time npm authorization (requires npm account 2FA once)

From an npm CLI v11.15+ session authenticated as the publishing account (npm account `otomasi`):

```bash
npm trust github @otomasi/office --repo rofiqcp/otomasi-office --file publish.yml --allow-publish --yes
```

The user may have to approve this one-time trust on npm website (2FA). The GitHub action uses OIDC only, with NO `NPM_TOKEN` secret. Configure within 2 days of first publish to validate the trust relationship. Alternatively via npm package **Settings → Trusted Publishing → GitHub Actions**, set owner `rofiqcp`, repository `otomasi-office`, workflow `publish.yml`, allow direct `npm publish`.

## Future releases (no npm OTP)

1. Update code and run `npm test` locally.
2. Increase the version in both `package.json` and `package-lock.json` (e.g. `npm version patch --no-git-tag-version`).
3. Commit and push changes to GitHub `main`.
4. Tag the exact release commit with `git tag vX.Y.Z` and push with `git push origin vX.Y.Z`.
5. GitHub Actions checks package version matches tag, runs `npm ci`, `npm test`, and publishes to npm using OIDC. No OTP or long-lived npm token.

The workflow is intentionally tag-gated. Ordinary pushes never publish to npm. No GitHub self-hosted runner; npm supports GitHub-hosted runners. For maximum account safety retain npm 2FA, use protected GitHub account, and avoid placing npm auth tokens in repository secrets.

**Already published and verified: 0.1.4. Never republish immutable npm version 0.1.4.** Bootstrap the trusted publisher with version 0.1.5 or later. Initial OIDC trust may expire if no successful publish occurs within 2 days; enable trust when ready for first release.
