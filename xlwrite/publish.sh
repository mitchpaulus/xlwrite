#!/bin/sh

dotnet publish -r linux-x64 -o publish/ -c Release -p:DebugType=None --no-self-contained -p:PublishSingleFile=true && \
cp publish/xlwrite "$LOCALBIN"/xlwrite
