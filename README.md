# EdgeSize
Launch Microsoft Edge at a specified size and position.

## Usage
If you want to display GitHub on the left side of the screen and Google on the right side, prepare the following INI file and run the application.
Please rewrite the paths in the INI file as appropriate for your environment.

~~~
[System]
Count=2
Sleep=500

[App1]
Path=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
Arg=--app=http://github.com/
ProcName=msedge
ClassName=Chrome_WidgetWin_1
NamePropertyLocal=アドレスと検索バー
NameProperty=address and search bar
Address=["http://"|"https://"].*github\.com.*
X=0
Y=0
Width=960
Height=1024

[App2]
Path=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
Arg=--app=http://www.google.com/
ProcName=msedge
ClassName=Chrome_WidgetWin_1
NamePropertyLocal=アドレスと検索バー
NameProperty=address and search bar
Address=["http://"|"https://"].*google\.com.*
X=960
Y=0
Width=960
Height=1024
~~~
