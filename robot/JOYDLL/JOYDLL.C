#include <windows.h>
#include <dos.h>
#include <stdio.h>

int FAR PASCAL ReadJoyA();
int FAR PASCAL ReadJoyB();
int FAR PASCAL ReadJoyC();
int FAR PASCAL ReadJoyD();

int FAR PASCAL LibMain(HANDLE hInstance, WORD wDataSeg,
	WORD wHeapSize, LPSTR lpszCmdLine)
{
if (wHeapSize != 0)
	UnlockData(0);
return 1;
}

int FAR PASCAL ReadJoyA()
{
	 union REGS inregs;
	 inregs.h.ah = 0x84;
	 inregs.x.dx = 1;
	 int86(0x15, &inregs, &inregs);
	 return inregs.x.ax;
}

int FAR PASCAL ReadJoyB()
{
	 union REGS inregs;
	 inregs.h.ah = 0x84;
	 inregs.x.dx = 1;
	 int86(0x15, &inregs, &inregs);
	 return inregs.x.bx;
}

int FAR PASCAL ReadJoyC()
{
	 union REGS inregs;
	 inregs.h.ah = 0x84;
	 inregs.x.dx = 1;
	 int86(0x15, &inregs, &inregs);
	 return inregs.x.cx;
}

int FAR PASCAL ReadJoyD()
{
	 union REGS inregs;
	 inregs.h.ah = 0x84;
	 inregs.x.dx = 1;
	 int86(0x15, &inregs, &inregs);
	 return inregs.x.dx;
}


