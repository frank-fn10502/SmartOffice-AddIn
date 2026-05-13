#!/usr/bin/env bash
set -euo pipefail

# 預設所有 SmartOffice 相關 repo 都以 develop 作為日常整合分支。
# 可用 SMARTOFFICE_BRANCH 覆寫，例如：
#   SMARTOFFICE_BRANCH=feature/foo ./scripts/pull-and-build.sh
BRANCH="${SMARTOFFICE_BRANCH:-develop}"
CONFIGURATION="${CONFIGURATION:-Debug}"
PLATFORM="${PLATFORM:-Any CPU}"

# 以腳本所在位置推導 SmartOffice repo，再用 sibling path 找 Hub repo。
# 若 Hub 放在其他位置，可用 SMARTOFFICE_HUB_ROOT 覆寫。
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SMARTOFFICE_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"
HUB_ROOT="${SMARTOFFICE_HUB_ROOT:-$(cd "${SMARTOFFICE_ROOT}/../SmartOffice.Hub" && pwd)}"

log() {
  printf '\n==> %s\n' "$*"
}

fail() {
  printf '\nERROR: %s\n' "$*" >&2
  exit 1
}

require_repo() {
  local repo_root="$1"
  local repo_name="$2"

  if [[ ! -d "${repo_root}/.git" ]]; then
    fail "${repo_name} 不是 git repository: ${repo_root}"
  fi
}

pull_repo() {
  local repo_root="$1"
  local repo_name="$2"

  log "Switch ${repo_name} to ${BRANCH}"
  if git -C "${repo_root}" show-ref --verify --quiet "refs/heads/${BRANCH}"; then
    git -C "${repo_root}" switch "${BRANCH}"
  elif git -C "${repo_root}" show-ref --verify --quiet "refs/remotes/origin/${BRANCH}"; then
    git -C "${repo_root}" switch --track -c "${BRANCH}" "origin/${BRANCH}"
  else
    fail "${repo_name} 找不到 local 或 origin/${BRANCH} branch，請先建立並 push ${BRANCH}。"
  fi

  # 使用 --ff-only 避免自動產生 merge commit；若分支分歧，請人工 rebase/merge 後再跑。
  log "Pull ${repo_name}/${BRANCH}"
  git -C "${repo_root}" pull --ff-only origin "${BRANCH}"
}

find_msbuild() {
  # 允許 Windows 主機明確指定 MSBuild 位置，避免不同 VS 安裝版本造成誤判。
  if [[ -n "${MSBUILD_EXE:-}" && -x "${MSBUILD_EXE}" ]]; then
    printf '%s\n' "${MSBUILD_EXE}"
    return 0
  fi

  # Git Bash / Developer Command Prompt 常見情境：MSBuild.exe 已在 PATH。
  if command -v MSBuild.exe >/dev/null 2>&1; then
    command -v MSBuild.exe
    return 0
  fi

  # 某些環境會提供 lowercase msbuild。
  if command -v msbuild >/dev/null 2>&1; then
    command -v msbuild
    return 0
  fi

  # Windows Git Bash 常見的 vswhere 路徑，用來尋找 Visual Studio 的 MSBuild。
  local vswhere='/c/Program Files (x86)/Microsoft Visual Studio/Installer/vswhere.exe'
  if [[ -x "${vswhere}" ]]; then
    local install_path
    install_path="$("${vswhere}" -latest -requires Microsoft.Component.MSBuild -property installationPath)"
    if [[ -n "${install_path}" ]]; then
      local candidate="${install_path}/MSBuild/Current/Bin/MSBuild.exe"
      if [[ -x "${candidate}" ]]; then
        printf '%s\n' "${candidate}"
        return 0
      fi
    fi
  fi

  return 1
}

require_repo "${SMARTOFFICE_ROOT}" "SmartOffice"
require_repo "${HUB_ROOT}" "SmartOffice.Hub"

pull_repo "${SMARTOFFICE_ROOT}" "SmartOffice"
pull_repo "${HUB_ROOT}" "SmartOffice.Hub"

# Hub 可在目前這台主機用 container build 驗證。
log "Build SmartOffice.Hub"
"${HUB_ROOT}/scripts/build-in-container.sh"

# Outlook Add-in 是 .NET Framework 4.8 VSTO 專案；完整 build 需要 Windows/VS/VSTO。
log "Build SmartOffice solution"
msbuild_path="$(find_msbuild)" || fail "找不到 MSBuild。SmartOffice Outlook Add-in 是 .NET Framework 4.8 VSTO 專案，需要 Windows 主機、Visual Studio/Build Tools 與 Office/VSTO 環境才能完整 build。"

# 有 nuget CLI 時先 restore packages；沒有時交給 MSBuild/Visual Studio 環境處理。
if command -v nuget >/dev/null 2>&1; then
  log "Restore NuGet packages"
  nuget restore "${SMARTOFFICE_ROOT}/SmartOffice.sln"
fi

"${msbuild_path}" "${SMARTOFFICE_ROOT}/SmartOffice.sln" \
  /m \
  /p:Configuration="${CONFIGURATION}" \
  /p:Platform="${PLATFORM}"

log "Pull and build completed"
