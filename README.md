# Plane 自宅サーバー導入手順書（Linux MintノートPC向け）

**作成日**: 2026年5月  
**対象環境**: 
- サーバー：Linux Mint（Intel Core i5 7th gen）
- クライアント：macOS + Androidスマホ
- 目的：外出先からも安全に使用

---

## サーバー情報
- **ローカルIP**: `192.168.10.110`
- **推奨ポート**: 8080
- **スペック**: Intel Core i5 7th / RAM 8GB以上推奨（満たしています）

---

## 1. Dockerインストール（Linux Mint）

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install curl git -y

# Docker公式インストール
curl -fsSL https://get.docker.com | sh -

# 権限設定
sudo usermod -aG docker $USER
newgrp docker

# 確認
docker compose version
