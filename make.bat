windres -o resources.o specan.rc
gcc -o specan.exe specan.c resources.o -lcomctl32 -mwindows -g