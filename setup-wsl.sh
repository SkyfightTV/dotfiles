#!/bin/bash

touch $HOME/.hushlogin
sudo apt update
sudo apt upgrade
sudo apt-add-repository ppa:fish-shell/release-3
sudo apt install fish
chsh -s /usr/bin/fish

Relaunch shell
# Copy config of fish
curl -sS https://starship.rs/install.sh | sh
# Copy starship config to ~/.config/starship.toml

sudo apt-get install ca-certificates curl gnupg
sudo mkdir -m 0755 -p /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
echo "deb [arch="$(dpkg --print-architecture)" signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu "$(. /etc/os-release && echo "$VERSION_CODENAME")" stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt-get update
sudo apt-get install docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
sudo usermod -aG docker $USER
echo -e "[boot]\nsystemd=true" | sudo tee -a /etc/wsl.conf
