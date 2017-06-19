all:
	cl -nologo -MD -EHsc -O2x -c .\main.cpp
	rc Version.rc
	link -out:PruneNetCfg.exe -opt:icf -opt:ref main.obj version.res ole32.lib

debug:
	cl -nologo -MDd -EHsc -Zi -c .\main.cpp
	rc Version.rc
	link -debug -out:PruneNetCfg.exe main.obj version.res ole32.lib
