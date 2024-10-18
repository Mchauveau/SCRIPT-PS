@echo off
REM Mappage d'un lecteur réseau temporaire
net use K: "\\srvfichiers\echanges\1 - DSI\SCRIPT TEST"

REM Exécution du script avec le lecteur réseau mappé
k:
cd "SetUserFTA"
SetUserFTA.exe
SetUserFTA  http ChromeHTML
SetUserFTA  https ChromeHTML
SetUserFTA  .htm ChromeHTML
SetUserFTA  .html ChromeHTML

REM Déconnexion du lecteur réseau après exécution
c:
net use K: /delete