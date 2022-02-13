#include <stdio.h>

extern int far ReadJoy(int far *, int far *, int far *, int far *);


main()
{
	int a,b,c,d;

	ReadJoy(&a,&b,&c,&d);

	printf("%d %d %d %d\n",a,b,c,d);

	return 0;
}
