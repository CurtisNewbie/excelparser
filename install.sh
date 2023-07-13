#!/bin/bash

git_repo="https://github.com/CurtisNewbie/excelparser"
app="excelparser"

CURR=$(pwd) \
	&& echo "Installing $app, previous working directory: $CURR" && echo \
	&& cd /tmp \
	&& echo "Cloing $app to /tmp" && echo \
	&& git clone "$git_repo" --depth 1 && (cd "$app" && python3 -m pip install .) \
	&& rm -rf "$app" \
	&& echo "/tmp/$app removed" \
	&& cd "$CURR"