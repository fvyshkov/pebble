@echo off
powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/fvyshkov/pebble/main/installer/install.ps1 | iex"
