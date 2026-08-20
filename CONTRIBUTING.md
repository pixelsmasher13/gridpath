# Contributing to GridPath

Thanks for your interest in contributing.

## Licensing of contributions

By submitting a pull request, issue, or any other contribution to this repository, you agree that your contribution is licensed to Heelix Technologies Inc. under the [Functional Source License, Version 1.1, with Apache 2.0 Future License](LICENSE) (FSL-1.1-Apache-2.0) on the same terms as the rest of the project.

You also represent that you have the right to license your contribution under those terms — i.e., the contribution is your original work, or if it's adapted from elsewhere, the source is FSL-1.1-Apache-2.0-compatible and you've credited it appropriately.

No separate CLA is required.

## Practical notes

- **Issues first for anything non-trivial.** Open an issue before sending a large PR so we can align on direction before you spend time. Tiny fixes (typos, obvious bugs) can come straight in.
- **Keep PRs focused.** One change per PR. If you're refactoring while fixing a bug, split it.
- **Don't introduce new dependencies without flagging.** Each dep is a long-term liability; we lean toward "use what we have."
- **Run the build before opening a PR.** `cargo check` on the Rust side and `npx tsc --noEmit` on the frontend should both be clean.

That's it.
