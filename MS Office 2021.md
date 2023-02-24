# MS Office 2021 Free Download and Install
>Follow The Given Steps

## Step 1:
Download for windows 64bit/32bit OS [Click Here](https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img)

## Step 2:
Open CMD in Run as Administrator 

## Step 3.1:
Get into the Office directory in cmd
```
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cd /d %ProgramFiles%\Microsoft Office\Office16
```

## Step 3.2
Install Office 2021 volume license
```
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```

## Step 4:
Activate your Office using the KMS key.
```
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:e8.us.to
cscript ospp.vbs /act
```

##### Congratulations, Office Activated 🥳
