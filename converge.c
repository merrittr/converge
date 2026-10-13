/*
File:		CONVERGE.C
Date:		April 10,1993.
Author:		Chris Milne
Purpose:	Convert multiple strain measurements at
		arbitrary angles to principal strain and angle
		of principal strain and determine stress.
Revision:	Feb 24 2002
		Modify calculation of epsilon1 and epsilon2
		Making this version 1.1                       */

#include <stdio.h>
#include <process.h>
#include <math.h>
#include <conio.h>
#include <string.h>


void intro(void);
FILE *getfilename( char *prompt, char *ext, char *root);
void headingprint(FILE *pntr);
void columntop(FILE *pntr);
void linedata(FILE *pntr);
void sig3fprintf(FILE *pntr, float val);

float version=1.1;			/* program version number */
char closure_station[20],looking[20];
char linetext[20];
float young,poisson,shear_modulus;
double angle[3];			/* angles from data file */
float sta_dist[3];			/* initial station distances */
float strainin[3];			/* input ea eb ec  */
float epsilon1,epsilon2,alpha;
float sigma1,sigma2;
float K,strain[3];		      	/* results, ex ey and gamma */
float p,den;

main (int argc, char *argv[])
{
    FILE *infileptr,*printptr,*spreadptr;
    char infilename[25];
    char indefault[]={"CONV.PRN"};
    char in_ext[]={".PRN"};
    char print_ext[]={".PPP"};
    char spread_ext[]={".SSS"};
    char printprompt[] = "printer";
    char spreadprompt[] = "spreadsheet";
    char *numbers[] = {"1st","2nd","3rd"};
    double mat[3][3],inmat[3][3];
    double radians,tmp;
    double rad_deg = 0.017453292519943;	/* radians / degree */
    int i,j,k,pivot,fstatus;

    intro();
    if (argc > 1) strcpy (infilename,argv[1]);
    else
    {
	printf("\nEnter input data filename: ");
	gets( infilename);
    }
    if( strlen(infilename) == 0) strcpy( infilename,indefault);
    else
    {
	if( strchr( infilename, '.' ) == 0) strcat( infilename, in_ext);
    }
    printf("\nInput filname is %s\n",infilename);

    if ((infileptr = fopen( infilename , "r")) == NULL)
    {
       printf ("\nCan't open file %s\n\a\a", infilename);
       exit (1);
    }
    *strchr( infilename, '.' ) = 0;	/* get root of infilename */
    strcpy (closure_station , infilename);
    if (argc > 2) strcpy ( looking, argv[2] );

/*Get the first line of the data file,
  3 angles and Young's modulus */
    fstatus = fscanf( infileptr,"%lf%lf%lf%f",
		      &angle[0],&angle[1],&angle[2],&young);
    printf("\nAngles         %10.1f%8.1f%8.1f",angle[0],angle[1],angle[2]);
    printf("\nYoung's modulus  %8.1f\n",young);

/* Get Poisson's ratio if other than 0.2 and
     determine the shear modulus */

    poisson = 0.2;
    printf("\nEnter Poisson's ratio if not 0.2 ");
    gets( linetext);
    if ( strlen (linetext))  sscanf( linetext , "%f", &poisson);

    printf("\nPoisson's ratio %4.2f",poisson);

    shear_modulus = 500 * young / ( 1 + poisson );

/* Open whichever output files are required */

    printptr = getfilename( printprompt,print_ext,infilename);
    spreadptr = getfilename( spreadprompt,spread_ext,infilename);


/* Make the stress matrix  " mat "   */

    for ( i=0 ; i<3 ; i++ )
    {
	radians = rad_deg * angle[i];
	tmp = Math.Cos( radians );
	mat[i][0] = tmp * tmp ;
	tmp = sin ( radians );
	mat[i][1] = tmp * tmp ;
	mat[i][2] = 0.5 * Math.Sin( 2 * radians ) ;
    }
    printf("\n matrix \n");
    for ( i=0 ; i<3 ; i++ )
    {
	for (j=0 ; j<3 ; j++ )  printf(" %12.5f",mat[i][j]);
	printf("\n");
    }

/* Initialize the future inverse matrix "inmat" to the identity matrix */

    for ( i=0 ; i<3 ; i++ )
    {
	for ( j=0 ; j<3 ; j++ ) inmat[i][j] = 0 ;
	inmat[i][i] = 1 ;
    }

/* Make the inverse matrix .  First do something like a pivot,
   to avoid getting a zero on the diagonal.  */
    pivot = 0;
    if ( Math.Abs( mat[0][0] ) < .000001) pivot = 1;
    else
    {
	if ( Math.Abs( mat[0][1] > .000001))
	{
	    tmp = mat[1][0] / mat[0][0];
	    if ( Math.Abs(tmp - mat[1][1]/mat[0][1]) < .000001) pivot =1;
	}
    }
    if ( pivot )
    {
	for ( j=0 ; j<3 ; j++ ) mat[0][j] += mat[2][j];
    }

/* Now the matrix inversion */

    for ( k=0 ; k<3 ; k++ )
    {
	if (( tmp =  mat[k][k] ) == 0 )
	    {
		printf ("\n****divide by zero\a\a" );
		exit (2);
	    }
	for ( j=0 ; j<3 ; j++ )
	{
	    mat[k][j] /= tmp , inmat[k][j] /= tmp;
	}
	for ( i=0 ; i<3 ; i++ )
	{
	    if ( i != k )
	    {
		tmp = mat[i][k];
		for ( j=0 ; j<3 ; j++ )
		{
		    mat[i][j]   = mat[i][j] -  mat[k][j] * tmp;
		    inmat[i][j] = inmat[i][j] - inmat[k][j] * tmp;
		}
	    }
	}
    }

/* Now compensate for the  sort of pivot. */
    if ( pivot )
    {
	for ( i=0 ; i<3 ; i++ ) inmat[i][2] += inmat[i][0];
    }

/* Show the result */

    printf("\n inverted matrix \n");
    for ( i=0 ; i<3 ; i++ )
    {
	for (j=0 ; j<3 ; j++ ) printf(" %12.5f",inmat[i][j]);
	printf("\n");
    }

/* If there has been a printer file specified, ask for
    the items to make the heading of the print file
    and send the heading to the file and the screen. */

    if (printptr != NULL )
    {
	printf("\n Enter the closure station");
	printf("\n if different from   %s  :",closure_station);
	gets(linetext);
	if (strlen(linetext)) strcpy (closure_station,linetext);
	for ( i=0 ; i<3 ; i++ )
	{
	    printf("\n Enter the %s initial station distance :",numbers[i]);
	    gets(linetext);
	    if ( strlen(linetext)) sscanf(linetext,"%f",&sta_dist[i]);
	    else sta_dist[i] = (float)(i+1);
	}
	fflush(stdin);
	if ( strlen(looking) == 0)
	{
	printf("\n Enter the orientation direction. Looking:");
	gets( looking );
	}
	printf("\n\n");

	headingprint( printptr);
	columntop (printptr);

	headingprint( stdout );

    }
    columntop(stdout);

/* Process each line of the input file */

    while ( fstatus != EOF )
    {
	fstatus = fscanf( infileptr, "%s%f%f%f",
			  linetext,&strainin[0],&strainin[1],&strainin[2]);
	if ( fstatus > 3)
	{
/* Determine epsilon1 and epsilon2 using;
	  epsilon x  =strain[0]
	  epsilon y  =strain[1]
	  gamma xy   =strain[2]  */

	  for ( i=0 ; i<3 ; i++ ) { strain[i] = 0; }
	  for ( i=0 ; i<3 ; i++ )
	  {
	      for ( j=0 ; j<3 ; j++ )
	      {
		 strain [i] +=  .000001 * ( strainin[j] * inmat[i][j] );
	      }
	  }

	  tmp = strain[0] - strain[1];
	  tmp = sqrt( strain[2] * strain[2] + tmp * tmp);
	  epsilon1 = 0.5 * (strain[0] + strain[1] + tmp);
	  epsilon2 = 0.5 * (strain[0] + strain[1] - tmp);

/* Determine alpha */

	  tmp = strain[0] - strain[1];
	  if ( Math.Abs( tmp ) > .0000001)
	  {
	    alpha = 0.5 * atan ( strain[2] / tmp) / rad_deg ;
	  }
	  if ( tmp < 0. )		/* left side of plane */
	  {
	    if ( strain[2] < 0. )	/* third quadrant */
	    {
	      if (tmp > -.00000011 ) alpha = -45.;
	      else alpha -= 90.;
	    }
	    else			/* second quadrant */
	    {
	      if (tmp > -.00000011 ) alpha = 45.;
	      else alpha += 90.;
	    }
	  }
	  else				/* right side of plane */
	  {
	    if ( strain[2] < 0. )	/* forth quadrant */
	    {
	      if ( tmp < .00000011 ) alpha = -45.;
	    }
	    else			/* first quadrant */
	    {
	      if ( tmp < .00000011 ) alpha = 45.;
	      if ( strain[2] == 0. ) alpha = 0.;	/* special case */
	    }
	  }

/* determine K, epsilon1 and epsilon2 */
	  p = 4 * (1-poisson) - 1;
	  den = epsilon1 * (1-p) - epsilon2 * (1+p);
	  K = (epsilon2 * (1-p) - epsilon1 * (1+p)) / den;
	  sigma2 = (epsilon1 * 4 * shear_modulus)/ ((1+K) - p * (1-K));
	  sigma1 = K * sigma2;
	  epsilon1 = 1000000. * epsilon1;
	  epsilon2 = 1000000. * epsilon2;

/* data to the screen */

	  linedata( stdout );

/* data to the print file if specified */

	  if ( printptr != NULL ) linedata( printptr);

/* data to the spreadsheet file if specified */

	  if ( spreadptr != NULL )
	  {
	  fprintf(spreadptr,"\n%s",linetext);
	  fprintf(spreadptr,"  %8.1f  %8.1f  %6.1f  %12.1f  %12.1f",
	     epsilon1,epsilon2,alpha,sigma1,sigma2);
	  }
	}
    }
    printf("\n");
    if( printptr != NULL )  fclose (printptr);
    if( spreadptr != NULL ) fclose (spreadptr);
    fclose ( infileptr );
    printf("\n");
    return (0);
}

/* For the spreadsheet output file and the printer output file.
   Asks for a filename and opens the file.  Returns a pointer
   to the file. */

FILE *getfilename( char *prompt,char *extension,char *root)
{
    FILE *pointer;
    char filename[25];
    fflush( stdin);
    printf("\n\nEnter %s output file",prompt);
    printf("\nif diffrent from %s%s",root,extension);
    printf("\nor enter N for NONE :");
    gets( filename );
    if ( strlen( filename ) == 0)
    {
	strcpy (filename,root);
	strcat (filename,extension);
	pointer = fopen(filename,"w");
    }
    else
    {
	if ( filename[0] == 'n' || filename[0] == 'N' ) pointer = NULL;
	else pointer = fopen(filename,"w");
    }
    return( pointer );
}

/* Produce the headings for the data in the printout file */


void headingprint(FILE *fptr)
{
    fprintf(fptr,"\t  Closure Station  %s \t  Looking  %s",
		  closure_station,looking);
    fprintf(fptr,"\n\nAngle CCW     Initial Station ");
    fprintf (fptr,"\nfrom horiz.    \t Distance");
    fprintf(fptr,"\nA= %5.1f\t a= %5.0f    E= %3.1f GPa   v= %3.2f   G= %3.0f MPa",
	     angle[0],sta_dist[0],young,poisson,shear_modulus);
    fprintf(fptr,"\nB= %5.1f\t b= %5.0f",angle[1],sta_dist[1]);
    fprintf(fptr,"\nC= %5.1f\t c= %5.0f",angle[2],sta_dist[2]);
    fprintf(fptr,"\n");
}

/* Used to produce the top of column headings for the result
   data. */


void columntop(FILE *fptr)
{
    fprintf(fptr,"\n               -------- micro strain -");
    fprintf(fptr,"-------");
    fprintf(fptr,"\n               ea     eb     ec     e1");
    fprintf(fptr,"     e2  alpha      sigma1      sigma2");
    fprintf(fptr,"\n");
}

/* Used to produce a line of results for the print file and for
    the screen display */

void linedata(FILE *fptr )
{
    fprintf(fptr,"\n%10s%7.0f%7.0f%7.0f",
	    linetext,strainin[0],strainin[1],strainin[2]);
    fprintf(fptr,"%7.0f%7.0f%7.1f  ",
       epsilon1,epsilon2,alpha);
    sig3fprintf(fptr, sigma1);
    sig3fprintf(fptr, sigma2);
}

/* Limit the significant figures of a printed number to 3, and
   display negative values in exponential format.  Used for
   sigma1 and sigma2 results.  */

void sig3fprintf(FILE *fptr, float n)
{
    float scale;

    if ( n >=0.0 )
    {
	if ( n< 99.5)
	{
	    if (n < 9.95 ) fprintf(fptr,"%12.2f",n);
	    else fprintf(fptr,"%11.1f ",n);
	}
	else
	{
	    scale = 1;
	    while( n > 999.5)
	    {
		n /=10;
		scale *=10;
	    }
	    n = scale * (float)(int)( n + 0.5);
	    fprintf(fptr,"%9.0f   ",n);
	}
    }
    else fprintf(fptr,"  %1.2E",n);
}

/* display introduction on the screen*/

void intro(void)
{
int i;
char *introduction[26]={" ",
"			     CONVERGENCE PROGRAM   ",
"                                                                              ",
"The program 'converge.exe' has been designed to assist in the interpretation  ",
"of closure station data.  For input, the program takes three values           ",
"of strain, a direction angle for each strain, Young's Modulus and Poisson's   ",
"Ratio.  A default value for Poisson's Ratio is set at .2, however, the user   ",
"should realize that this value can have a significant effect on results.      ",
"                                                                              ",
"Program output consists of principal strain values and orientation, as well   ",
"as principal stresses.  The theory used for obtaining principal strains and   ",
"orientation is based on a matrix inversion and solution taken from            ",
"\"Introduction to Rock Mechanics\" by Goodman (1989).  The resolution of principal",
"stresses from strains applies the Kirsch solution taken from \"Rock Mechanics  ",
"for Underground Mining\" by Brady & Brown, (1992).  This solution assumes a    ",
"circular tunnel geometry and an elastic rock mass response.  When the elastic ",
"response of the rock mass appears to have been exceeded, the orientation and  ",
"magnitude of principal strains may still provide useful data.                 ",
"                                                                              ",
"This program has been developed by Milne & Milne Engineering in conjunction   ",
"with the UBC Mining-Geotechnical Group.  Assistance has been provided by the  ",
"Noranda Technology Centre and Energy Mines and Resources.  If you find this   ",
"program of use, please send 12 chocolate chip cookies (no nuts please) to     ",
"Milne & Milne Engineering, 11067 129 Street, Surrey BC  V3T 3J3               ",
"press ENTER",
" "};
  printf("\n%s",&introduction[1][0]);
  printf(" (Version %3.1f)",version);
  for ( i=2 ; i<25 ; i++ )
  {
  printf("\n%s",&introduction[i][0]);
  }
  fflush(stdin);
  gets (&introduction[0][0]);
  printf("\n\n");
}
