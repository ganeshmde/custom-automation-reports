@ECHO off
SET PROJECT_NAME=CustomExtentReport
dotnet build %PROJECT_NAME%.sln
for /l %%x in (1, 1, 5) do (echo:)
call bin\Debug\net8.0\%PROJECT_NAME%.exe
pause