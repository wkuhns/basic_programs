!*******************************************************************
!
!   Program to calculate and plot drive wheel torques - see DRAG.DOC
!   for detailed documentation.
!
!*******************************************************************

MAIN:
    gosub make_tables

    open "goals.dat" for input as #1%
!	pages, 
!	frame mass, 
!	number of springs, 
!	max spring ext in inches,
! 	tire diameter,
!	string to spring ratio
!	drive train efficiency
!	mu
!	final drive ratio
    input #1%, pages,fmass,nsprings,maxext,tire_dia,pulley_ratio,eff,mu,fdr
    mass = fmass + (0.13 * nsprings)
    Print "Generating ";pages;"pages"
    print "Vehicle mass = ";mass
    print "Powered by ";nsprings;"springs at ";maxext;"extension"
    print "Tire diameter of ";tire_dia;"inches"
    print "Drive train efficiency of";eff
    print "Mu =",mu
    open "logfile.dat" for output as #2%

    for xx=1 to pages
 	run$ = str$(xx)
	gosub init_tables
	print "Set tables, run ";run$
	gosub calc_winds
	print "Calc winds, run ";run$
	gosub calc_torques
	print "Calc torques, run ";run$
	gosub graph_torques
	print "Graphed run ";run$

	!*******************************************************************
	!   Format report header data
	!*******************************************************************

	ext$ = str$(maxext)+"/"+str$(actext)
	gosub write_graph			! Write output to file
    next xx

GOTO END_DRAG

!
!*******************************************************************
!
!   Create and initialize various tables used in program
!
!*******************************************************************

MAKE_TABLES:

    true = -1
    false = 0

    dim dias(16)		    ! Lookup table of spindle diameters
    dim turns(16)		    ! Turns per diameter (wind schedule)
    dim runs(16)		    ! Inches ground travel per cord inch
    dim dists(16)		    ! Ground travel for this diameter
    dim inches(16)		    ! Cum. inches at end of segment
    
    dim torques(200)		    ! Drive torques at 2" increments
    dim screen$(80,132)		    ! Graph image
    dim chart$(2,132)		    ! Windup chart strip

    dias(1) = .2688		    ! Initialize with actual diameters
    dias(2) = .2889
    dias(3) = .3106
    dias(4) = .3339
    dias(5) = .3589
    dias(6) = .3858
    dias(7) = .4148
    dias(8) = .4459
    dias(9) = .4793
    dias(10) = .5153
    dias(11) = .5539
    dias(12) = .5954
    dias(13) = .6401
    dias(14) = .6881
    dias(15) = .7397
    dias(16) = .7952

RETURN

INIT_TABLES:

    for x=0 to 132
	chart$(1,x)=" "
	chart$(2,x)=" "
    next x
    chart$(1,0)="|"
    chart$(2,0)="^"
   
   seq$ = ""

RETURN

!
!*******************************************************************
!
!   Calculate optimized wind schedule
!
!*******************************************************************

CALC_WINDS:

    for x=1 to 16
	turns(x) = 0
    next x

    !***************************************************************
    !	Various vehicle parameters in english / metric hybrid
    !***************************************************************

    fgoal = mass * mu				! force goal
    sconst = 3.065 * nsprings			! in pounds per foot
    sconst = sconst / pulley_ratio
    offset = (nsprings * 8.6) / pulley_ratio	! in pounds

    !***************************************************************
    !	Initial condition is vehicle with springs fully extended
    !	and cord on smallest spindle diameter
    !***************************************************************

    travel = 0					! ground travel in inches
    seg = 1					! seg 1 is smallest

    ext = maxext
    torque = ((ext*sconst)+offset)*(dias(seg)/2)
    dforce = torque / (tire_dia * fdr / 2)
!
    !***************************************************************
    !	If torque is too low, move to larger spindle diameters. Each
    !	Diameter is 7.5% more torque, so we start 4% high.
    !***************************************************************

    while (dforce < (fgoal*1.04))
	seg = seg+1
	torque = ((ext*sconst)+offset)*(dias(seg)/2)
        dforce = torque / (tire_dia * fdr / 2)
        print "Had to move first winding to segment";seg
    next

    !***************************************************************
    !	We now have at least enough torque. Unwind .05 in at a time 
    !	until we are OK.
    !***************************************************************

    while (dforce > (fgoal*1.04))
	ext = ext - .05
	torque = ((ext*sconst)+offset)*(dias(seg)/2)
        dforce = torque / (tire_dia * fdr / 2)
    next
    print "Unwound to ";ext
    actext = ext

    !***************************************************************
    !	Now we do the work. CALC_SEG determines the best spindle 
    !	segment by balancing torque errors.
    !	Each increment is 1/4 tire rotation.
    !***************************************************************

    finished = false

    while (finished = false)
	gosub calc_seg			! Call twice to skip a segment
	gosub calc_seg			! if needed
	turns(seg) = turns(seg) +.25
	ext = ext - (dias(seg)*3.1416*.25)/pulley_ratio
	travel = travel + 3.1416 * tire_dia * fdr *.25
    next

RETURN

!*******************************************************************
!
!   Given a wind schedule, calculate the actual torques delivered
!   to the rear wheels
!
!*******************************************************************

CALC_TORQUES:

    !***************************************************************
    !	DISTS() is the ground distance traveled on each spindle
    !	diameter.
    !***************************************************************

    travel = 0
    for x=1 to 16
        dists(x) = 3.1416 * tire_dia * fdr * turns(x)	    ! Calculate DISTS()
	travel = travel + dists(x)
    next x
    print "Total powered distance = ";travel;" inches"
    ext = 0

    for x=16 to 1 step -1
        ext = ext + turns(x) * dias(x) * 3.1416 / pulley_ratio
    next x
    calcext = ext + .33

    !***************************************************************
    !   Calculate the torque delivered to the wheels at 2" increments of 
    !   ground travel. Torque is a function of spring extension and spindle
    !   diameter.
    !***************************************************************

    travel = 0
    dist = 0
    seg = 0

    rot_incr = 1 / (tire_dia * fdr * 3.1416)

    ext = actext		    ! Actual extension is in inches

    !***************************************************************
    !	MASS, V_I, and E_T are used in theoretical performance
    !	calculations. See drag.doc for details.
    !***************************************************************

    v_i = 0
    e_t = 0

    for inch=1 to 400
	x=int(inch/2)
        while ((dist<1) and (seg<16))   ! Accumulate at least 1" of travel
 	    seg = seg + 1		! by traversing empty segments
   	    dist = dist + dists(seg)    ! until you find enough DISTance
	    gosub chart
        next
				    ! If you have any cord left, calculate
				    ! torque as a function of spring force
				    ! and spindle diameter.

        if dist >= 1 then
	    torques(x) = (ext * sconst + offset) * (dias(seg)/2) 
  	    dist = dist - 1	    ! Decrement remaining distance
	    ext = ext - (dias(seg)*3.1416 * rot_incr)/pulley_ratio
	else
	    torques(x) = 0
	end if

    !***************************************************************
    !   Theoretical calculations mentioned above.
    !***************************************************************

	force = torques(x) / (tire_dia * fdr /2)
	force = force * eff

	acc = (force/mass) * 32

	if acc > 0 then
	    d_t = (sqrt(v_i*v_i+2*acc/12)-v_i)/acc
	else
	    d_t = 1/v_i/12
	end if
	e_t = e_t + d_t

	v_i = v_i + acc*d_t

	if inch=15 then
	    a_a = 2.5 / (e_t * e_t)
	    aa$ = format$(a_a,"##.##")
	    ta$ = format$(e_t,"#.###")
	    va$ = format$(v_i,"##.##")
	    eta = e_t
	    vel = v_i
	    print "x";x;" torque";torques(x);" force";force;" acc";acc;" ext";ext
	end if

	if inch=60 then
	    etb = e_t - eta
	    a_b = 2*(3.75-(vel*etb)) / (etb * etb)
	    ab$ = format$(a_b,"##.##")
	    tb$ = format$(etb,"#.###")
	    vb$ = format$(v_i,"##.##")
	    vel = v_i
	    print "x";x;" torque";torques(x);" force";force;" acc";acc;" ext";ext
	end if

	if inch=135 then
	    etc = e_t - etb - eta
	    a_c = 2*(6.25-(vel*etc)) / (etc * etc)
	    ac$ = format$(a_c,"##.##")
	    tc$ = format$(etc,"#.###")
	    vc$ = format$(v_i,"##.##")
	    vel = v_i
	    print "x";x;" torque";torques(x);" force";force;" acc";acc;" ext";ext
	end if

	if inch=240 then
	    etd = e_t - etc - etb -eta
	    a_d = 2*(8.75-(vel*etd)) / (etd * etd)
	    ad$ = format$(a_d,"##.##")
	    td$ = format$(etd,"#.###")
	    vd$ = format$(v_i,"##.##")
	    net_acc= 40 / (e_t * e_t)
	    net_acc$ = format$(net_acc,"##.##")
	    et$ = format$(e_t,"#.###")
	    print "x";x;" torque";torques(x);" force";force;" acc";acc;" ext";ext
            print "Elapsed time:";e_t
	end if
    next inch
RETURN
!
!*******************************************************************
!
!   Create output graph
!
!*******************************************************************

GRAPH_TORQUES:

    for row=0 to 80			    ! Clear graph
        for col=1 to 132
	    screen$(row,col)=" "
        next col
    next row

    !***************************************************************
    !	Paint frame
    !***************************************************************

    for col=4 to 130 step 6
	screen$(9,col) = "|"
	screen$(61,col) = "|"
	xlabel = mod(((col-4)/6),10)
	screen$(8,col) = str$(xlabel)
	screen$(62,col) = str$(xlabel)
    next col

    for row=10 to 60 step 10
	for col=3 to 132
	    screen$(row,col) = "-"
	next col
    next row

    screen$(60,1)="1"
    screen$(60,2)="2"

    screen$(50,1)="1"
    screen$(50,2)="1"

    screen$(40,1)="1"
    screen$(40,2)="0"

    screen$(30,1)="."
    screen$(30,2)="9"

    screen$(20,1)="."
    screen$(20,2)="8"

    screen$(10,1)="."
    screen$(10,2)="7"

!
    for x=1 to 128			    ! place torque points in graph
	acc =  torques(x) * eff / ((tire_dia * fdr /2) * mass)
        xpos = int(acc * 100 -70) + 5
        if (xpos < 0) then
	    xpos = 0
	end if
	if xpos > 80 then
	    xpos = 80
	end if
	screen$(xpos,x+4) = "*"
    next x

    !***************************************************************
    !	Place theoretical data in graph
    !***************************************************************

    for x = 1 to 5
	screen$(13,8+x)=mid$(ta$,x,1)
	screen$(12,8+x)=mid$(aa$,x,1)
	screen$(11,8+x)=mid$(va$,x,1)

	screen$(13,31+x)=mid$(tb$,x,1)
	screen$(12,31+x)=mid$(ab$,x,1)
	screen$(11,31+x)=mid$(vb$,x,1)

	screen$(13,68+x)=mid$(tc$,x,1)
	screen$(12,68+x)=mid$(ac$,x,1)
	screen$(11,68+x)=mid$(vc$,x,1)

	screen$(13,121+x)=mid$(td$,x,1)
	screen$(12,121+x)=mid$(ad$,x,1)
	screen$(11,121+x)=mid$(vd$,x,1)
    next x
    
RETURN

!
!*******************************************************************
!
!   Write graph to disk
!
!*******************************************************************

WRITE_GRAPH:

    margin #2%, 132
    print #2%, run$
    print #2%
    print #2%
    print #2%
    print #2%
    print #2%,"Run: ";run$;
    print #2%,"   Mass: ";mass;
    print #2%,"   Tire: ";tire_dia;
    print #2%,"   Ratio: ";pulley_ratio;
    print #2%,"   Springs: ";nsprings;
    print #2%,"   Ext: ";ext$;
    print #2%,"   Acc: ";net_acc$;
    print #2%,"   Eff: ";eff;
    print #2%,"   Time: ";et$;
    print #2%,
    print #2%,
    for row=62 to 8 step -1
	for col=1 to 132
	    print #2%, screen$(row,col);
	next col
	print #2%,
    next row
    print #2%,

    for x=0 to 130
	print #2%, chart$(1,x);
    next x
    print #2%,
    for x=0 to 130
	print #2%, chart$(2,x);
    next x

    print #2%, chr$(12)

RETURN

!
!*******************************************************************
!
!   Create windup chart at bottom of graph
!
!*******************************************************************

CHART:

    ptr_cpi = 13.62		    ! char/in for ln03r in landscape

    cell = ext * ptr_cpi
    cell = int(cell + .5)

    if cell < 0 then
	cell = 0
    end if

    chart$(1,cell) = "|"
    chart$(2,cell) = chr$(seg+64)

RETURN

!*******************************************************************
!
!   Determine ideal segment based on minimizing torque errors
!
!*******************************************************************

CALC_SEG:

    if (turns(16) = 3) or (ext <= 0) then
	finished = true
    end if

    if (finished = true) or (seg = 16) then
	goto done_calc_seg
    end if

    !***************************************************************
    !	GOAL1 is desired torque at start of this increment. GOAL2 is 
    !	desired torque at end of this increment.
    !***************************************************************

    goal1 = (fgoal*fdr*tire_dia/2)/eff
    goal2 = goal1

!
    !***************************************************************
    !	ERROR1 is the maximum torque error if we stay on the current
    !	segment. ERROR2 is the max error if we move to the next.
    !***************************************************************

    torque1 = ((ext * sconst)+offset)*(dias(seg)/2)
    ext2 = ext - (dias(seg)*3.1416*.25/pulley_ratio)
    torque2 = ((ext2 * sconst)+offset)*(dias(seg)/2)
    error1 = abs(torque1-goal1)+abs(torque2-goal2)

    torque3 = ((ext * sconst)+offset)*(dias(seg+1)/2)
    ext2 = ext - (dias(seg+1)*3.1416*.25/pulley_ratio)
    torque4 = ((ext2 * sconst)+offset)*(dias(seg+1)/2)
    error2 = abs(torque3-goal1)+abs(torque4-goal2)

    if (error1>error2) then
	seg = seg + 1
    end if

DONE_CALC_SEG:
     
RETURN

END_DRAG:


