how to remove / start from local storage
crtl shift i
application tab
storage / local storage --> clear
indexeddb, session storage, cookies --> clear if everything needs to be gone


after failed npm run dist:
close cursor
open terminal as adminstrator
delete dist file
rmdir C:\Users\mehek\AppData\Local\electron-builder\Cache\winCodeSign -Recurse -Force
cd D:\projects\email-wrapper\butter-mail
npm run dist (again)