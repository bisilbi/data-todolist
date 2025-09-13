:: sebelum menjalankan file ini, store password dulu https://stackoverflow.com/questions/33919769/git-bat-password-not-working-in-cmd
rem git config --global credential.helper cache
rem git config --global credential.helper wincred
start cmd /c "git add . && git commit -m "update" && git push -u origin main && pause"
rem git config --global http.postBuffer 157286400