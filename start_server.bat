@echo off
echo Checking and installing PM2 if needed...
call npm install pm2

echo Starting server with PM2...
:: Restart if already exists, else start
call npx pm2 restart anti_test_server || call npx pm2 start server/server.js --name "anti_test_server"

echo Saving PM2 process list...
call npx pm2 save
echo.
echo =========================================================
echo Server is now running in the background via PM2.
echo It will automatically restart if it crashes.
echo.
echo Commands you can use:
echo - View status: npx pm2 status
echo - View logs:   npx pm2 logs anti_test_server
echo - Stop server: npx pm2 stop anti_test_server
echo =========================================================
pause
