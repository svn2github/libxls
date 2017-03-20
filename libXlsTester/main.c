/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * This file is part of libxls -- A multiplatform, C/C++ library
 * for parsing Excel(TM) files.
 *
 * Redistribution and use in source and binary forms, with or without modification, are
 * permitted provided that the following conditions are met:
 *
 *    1. Redistributions of source code must retain the above copyright notice, this list of
 *       conditions and the following disclaimer.
 *
 *    2. Redistributions in binary form must reproduce the above copyright notice, this list
 *       of conditions and the following disclaimer in the documentation and/or other materials
 *       provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY David Hoerl ''AS IS'' AND ANY EXPRESS OR IMPLIED
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
 * FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL David Hoerl OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 * SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 * ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 * Copyright 2008-2014 David Hoerl
 *
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <assert.h>
//#include <time.h>

#include <libxls/xls.h>

#include "xlsformula.h"

int main(int argc, char *argv[])
{
    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    struct st_row_data* row;
    WORD t,tt;
    unsigned int i;

int debug = 10;
xls(debug);	// set debug to 0
xls_set_formula_hander(dump_formula);

	if(argc != 2) {
		printf("Need file arg\n");
		exit(0);
	}

//system("pwd");
//system("ls -l David.xls");

	printf("Open file: %s\n", argv[1]);

    pWB=xls_open(argv[1],"UTF-8");	// man iconv_open for list of possible values - not sure which ones Excel uses	// ASCII

#if 1
xlsSummaryInfo *si = xls_summaryInfo(pWB);
printf("title=%s\n", si->title);
printf("subject=%s\n", si->subject);
printf("author=%s\n", si->author);
printf("keywords=%s\n", si->keywords);
printf("comment=%s\n", si->comment);
printf("lastAuthor=%s\n", si->lastAuthor);
printf("appName=%s\n", si->appName);
printf("category=%s\n", si->category);
printf("manager=%s\n", si->manager);
printf("company=%s\n", si->company);

xls_close_summaryInfo(si);
#endif
	//exit(0) ;
    if (pWB!=NULL)
    {
		assert(pWB->sheets.count);
		for (i=0;i<pWB->sheets.count;i++) {
			printf("\n---------------------------------------------\n");
            printf("  Sheet[%i] (%s) pos=%i\n",i, pWB->sheets.sheet[i].name, pWB->sheets.sheet[i].filepos);
			printf("---------------------------------------------\n");

			pWS=xls_getWorkSheet(pWB,i);

			xls_parseWorkSheet(pWS);

			printf("Count of rows: %i\n",pWS->rows.lastrow + 1);
			printf("Max col: %i\n",pWS->rows.lastcol);

#if 1
			for (t=0;t<=pWS->rows.lastrow;t++)
			{
				row=&pWS->rows.row[t];
				if(debug) xls_showROW(row);
				for (tt=0;tt<=pWS->rows.lastcol;tt++)
				{
					xlsCell	*cell;
					
					cell = &row->cells.cell[tt];

                    if(cell->id == 0x201) continue;
					printf("===================\n");
					printf("cell_id=%3.3x row=%d col=%d\n", cell->id, cell->row, cell->col);
					xls_showCell(cell);
					
					if(cell->id == 0x06) { // formula
					  if(cell->l == 0) { // its a number
						  printf("FORMULA: CELL NUMBER: %g %s\n", cell->d, cell->str);
					  } else {
						  if(!strcmp(cell->str, "bool"))  printf("Bool=%d", (int)cell->d);
						  if(!strcmp(cell->str, "error")) printf("ERROR\n");
						  else printf("FORMULA STRING: %s\n", cell->str);
					  }
					}
					//xls_showCell(&row->cells.cell[tt]);
				}
			}
#endif
			xls_close_WS(pWS);
		}
		//xls_showBookInfo(pWB);
        xls_close_WB(pWB);
    } else {
		printf("pWB == NULL\n");
	}
printf("Bye\n");
    return 0;
}


#if 0
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <assert.h>
//#include <time.h>

#include <libxls/xls.h>

int main(int argc, char *argv[])
{
    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    struct st_row_data* row;
    WORD t,tt;
    int i;

	xls(10);
	//xls_debug = 10; // 10;

	if(argc != 2) {
		printf("Need file arg\n");
		exit(0);
	}

    pWB=xls_open(argv[1],"UTF-8");	// man iconv_open for list of possible values - not sure which ones Excel uses

	//exit(0) ;

    if (pWB!=NULL)
    {
		assert(pWB->sheets.count);
        for (i=0;i<pWB->sheets.count;i++) {
            printf("Sheet[%i] (%s) pos=%i\n",i, pWB->sheets.sheet[i].name, pWB->sheets.sheet[i].filepos);

			pWS=xls_getWorkSheet(pWB,i);

			xls_parseWorkSheet(pWS);

			printf("Count of rows: %i\n",pWS->rows.lastrow + 1);
			printf("Max col: %i\n",pWS->rows.lastcol);
			
			for (t=0;t<=pWS->rows.lastrow;t++)
			{
				row=&pWS->rows.row[t];
				if(xls_debug) xls_showROW(row);
				for (tt=0;tt<=pWS->rows.lastcol;tt++)
				{
					//printf("cell=%x\n", &row->cells.cell[tt]);
					//xls_showCell(&row->cells.cell[tt]);
				}
			}
			
		}
		//xls_showBookInfo(pWB);
    } else {
		printf("pWB == NULL\n");
	}
    return 0;
}
#endif
