#!/bin/sh
# Release the version in app/Info.plist: build the zips, tag, draft a GitHub Release, smoke-test it on Apple silicon,
# publish it, then point the Homebrew cask at it.
set -e
ROOT=$(cd "$(dirname "$0")/.." && pwd)
REPO=ElvisLabs/zhibi
TAP="$(brew --repository)/Library/Taps/elvislabs/homebrew-tap"
cd "$ROOT"
ghgit() { git -c credential.helper= -c 'credential.helper=!gh auth git-credential' "$@"; }   # push as the gh account
[ -z "$(git status --porcelain)" ] || { echo "有未提交的改动，先提交再发版" >&2; exit 1; }
V=$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' app/Info.plist)
git rev-parse -q --verify "refs/tags/v$V" >/dev/null && { echo "v$V 已经发过，先改 app/Info.plist 的版本号" >&2; exit 1; }
app/build.sh --release
SHA_ARM=$(awk '/arm64\.zip/ {print $1}' dist/SHA256SUMS)
SHA_INTEL=$(awk '/x86_64\.zip/ {print $1}' dist/SHA256SUMS)
git tag "v$V"
ghgit push -q origin HEAD "v$V"
# Draft first: the smoke workflow runs the cross-built arm64 zip on Apple silicon before anyone can install it.
gh release create "v$V" dist/Zhibi-"$V"-*.zip dist/SHA256SUMS --repo "$REPO" --draft --title "执笔 $V" --notes "$(cat <<NOTES
安装：\`brew install --cask elvislabs/tap/zhibi\`，升级：\`brew upgrade --cask zhibi\`。

依赖和首次设置见 [README](https://github.com/$REPO#readme)。直接下载 zip 的话，应用没有经过 Apple 公证，解压后先运行 \`xattr -dr com.apple.quarantine Zhibi.app\` 再打开。Apple 芯片下载 arm64，Intel 下载 x86_64。
NOTES
)"
gh workflow run smoke.yml --repo "$REPO" -f tag="v$V"
sleep 8
RUN=$(gh run list --repo "$REPO" --workflow smoke.yml --limit 1 --json databaseId --jq '.[0].databaseId')
gh run watch "$RUN" --repo "$REPO" --exit-status > /dev/null \
  || { echo "冒烟测试没过，v$V 还是草稿：gh run view $RUN --repo $REPO --log-failed" >&2; exit 1; }
gh release edit "v$V" --repo "$REPO" --draft=false > /dev/null
sed -i '' -e "s/^  version \".*\"/  version \"$V\"/" -e "s/^  sha256 arm: *\".*\",/  sha256 arm:   \"$SHA_ARM\",/" \
  -e "s/^         intel: \".*\"/         intel: \"$SHA_INTEL\"/" "$TAP/Casks/zhibi.rb"
git -C "$TAP" add Casks/zhibi.rb
git -C "$TAP" commit -qm "Update zhibi to v$V"
ghgit -C "$TAP" push -q
echo "v$V 已发布：brew upgrade --cask zhibi"
