# Contributing to GridPath

Thanks for your interest in contributing.

## Licensing of contributions

Contributions are licensed under the license of the files they change, on
the same terms as the rest of that code ("inbound = outbound"):

- **Apache-2.0** for the engine, the saver, the `gridpath` package, the
  examples and the evals (most of the repository; see [NOTICE](NOTICE)).
- **FSL-1.1-Apache-2.0** for the desktop app (`src/`, `src-tauri/`) and the
  hosted services, which carry their own LICENSE files.

Sign off every commit (`git commit -s`). The `Signed-off-by:` line certifies
the [Developer Certificate of Origin 1.1](https://developercertificate.org/):
that you wrote the change or otherwise have the right to submit it under the
project's license. No separate CLA is required.

"GridPath" and the logo are trademarks; see [TRADEMARKS.md](TRADEMARKS.md).

## Practical notes

- **Issues first for anything non-trivial.** Open an issue before sending a large PR so we can align on direction before you spend time. Tiny fixes (typos, obvious bugs) can come straight in.
- **Keep PRs focused.** One change per PR. If you're refactoring while fixing a bug, split it.
- **Don't introduce new dependencies without flagging.** Each dep is a long-term liability; we lean toward "use what we have."
- **Run the build before opening a PR.** `cargo check` on the Rust side and `npx tsc --noEmit` on the frontend should both be clean.

That's it.
