DECLARE SUB SystemOutputs ()
DECLARE SUB PresentStatus ()
DECLARE SUB PresentFaults ()
DECLARE SUB FaultHistory ()
DECLARE SUB SelfTest ()
DECLARE SUB choice5 ()
DECLARE SUB choice6 ()
DECLARE SUB choice7 ()
    DIM m1string$(2, 8)
    mchoice = 0
    mchoices = 5
    m1string$(0, 0) = "System"
    m1string$(1, 0) = "Outputs?"
    m1string$(0, 1) = "Present"
    m1string$(1, 1) = "Status?"
    m1string$(0, 2) = "Present"
    m1string$(1, 2) = "Faults?"
    m1string$(0, 3) = "Fault"
    m1string$(1, 3) = "History?"
    m1string$(0, 4) = "Self"
    m1string$(1, 4) = "Test?"
    m1string$(0, 5) = "Choice F"
    m1string$(1, 5) = "Choice F line 2"
    m1string$(0, 6) = "Choice G"
    m1string$(1, 6) = "Choice G line 2"
    m1string$(0, 7) = "Choice H"
    m1string$(1, 7) = "Choice H line 2"


    inchar$ = "x"
    mchoice = 0
    WHILE inchar$ <> "q"
        SELECT CASE inchar$
        CASE IS = "n"
            mchoice = (mchoice + 1) MOD (mchoices)

        CASE IS = "y"
            SELECT CASE mchoice
                CASE IS = 0
                    CALL SystemOutputs
                CASE IS = 1
                    CALL PresentStatus
                CASE IS = 2
                    CALL PresentFaults
                CASE IS = 3
                    CALL FaultHistory
                CASE IS = 4
                    CALL SelfTest
                CASE IS = 5
                    CALL choice5
                CASE IS = 6
                    CALL choice6
                CASE IS = 7
                    CALL choice7
            END SELECT
        END SELECT
        CLS
        LOCATE 20, 32
        PRINT m1string$(0, mchoice)
        LOCATE 21, 32
        PRINT m1string$(1, mchoice)
        DO
        inchar$ = INKEY$
        LOOP WHILE inchar$ = ""
    WEND
STOP

SUB choice5 :
    CLS
    PRINT "Inside F"

END SUB

SUB choice6 :
    CLS
    PRINT "Inside G"

END SUB

SUB choice7 :
    CLS
    PRINT "Inside H"

END SUB

SUB FaultHistory :
    DIM msg$(2, 8)
    DIM inchar$
    DIM mchoice
    DIM mchoices
    ' $STATIC
    DIM msg$(2, 8)
    STATIC inchar2$
    STATIC schoice
    STATIC schoices

    PRINT "Fault History Count = 11"
    msg$(0, 0) = "FAULT HISTORY:"
    msg$(0, 0) = "## FAULTS"

    ' Last Leg Flown = # 75
    msg$(0, 0) = "LAST FLIGHT"
    msg$(0, 0) = "LEG NUMBER: 75"

    msg$(0, 0) = "FAULTS FOR"
    msg$(0, 0) = "LEG NUMBER 75?"
    ' (yes/no, etc)

    ' Fault #1 = Leg 75 (latest leg) Nose Gear Up Limit Sensor Open
    ' Circuit on ground
    msg$(0, 0) = "NOSE GEAR UPLMT"
    msg$(0, 0) = "SENSOR OPEN   G"

    ' Fault #2 = Leg 75 (latest leg) Nose Gear Down Limit Sensor Rigged
    ' Too Close on ground
    msg$(0, 0) = "NOSE GEAR DNLMT"
    msg$(0, 0) = "SMALL GAP .04 G"

    ' Fault #3 = Leg 75 (latest leg) Left Main Gear Down Lock
    ' Conditioner Offset (code F1) in air, requiring board A1 to be
    ' replaced
    msg$(0, 0) = "LEFT GEAR DNLMT"
    msg$(0, 0) = "BOARD A1 F1   A"

    msg$(0, 0) = "FAULTS FOR "
    msg$(0, 0) = "LEG NUMBER 74?"
    ' (yes/no, etc)

    ' Fault #4 = Leg 74 Nose Gear Up Limit Sensor Open Circuit in air
    msg$(0, 0) = "NOSE GEAR UPLMT"
    msg$(0, 0) = "SENSOR OPEN   A"

    ' Fault #5 = Leg 74 Nose Gear Up Limit Sensor Too Close in air,
    ' Rigging Required
    msg$(0, 0) = "NOSE GEAR UPLMT "
    msg$(0, 0) = "SMALL GAP .01 A"

    ' Fault #6 = Leg 74 Nose Gear Down Limit Sensor Rigged Too Close on
    ' ground
    msg$(0, 0) = "NOSE GEAR DNLMT"
    msg$(0, 0) = "SMALL GAP .04 G"

    ' Fault #7 = Leg 73 Nose Gear Up Limit Sensor Too Close in air,
    ' Rigging Required
    msg$(0, 0) = "NOSE GEAR UPLMT"
    msg$(0, 0) = "SMALL GAP .02 A"

    msg$(0, 0) = "FAULTS FOR "
    msg$(0, 0) = "LEG NUMBER 73?"
    ' (yes/no, etc)

    ' Fault #8 = Leg 73 Nose Gear Down Limit Sensor Rigged Too Close on
    ' ground
    msg$(0, 0) = "NOSE GEAR DNLMT"
    msg$(0, 0) = "SMALL GAP .04 G"

    msg$(0, 0) = "FAULTS FOR "
    msg$(0, 0) = "LEG NUMBER 72?"
    ' (yes/no, etc)

    ' Fault #9 = Leg 72 Nose Gear Up Limit Sensor Too Close, Rigging
    ' Required in air
    msg$(0, 0) = "NOSE GEAR UPLMT"
    msg$(0, 0) = "SMALL GAP .02 A"

    ' Fault #10 = Leg 72 Nose Gear Down Limit Sensor Rigged Too Close
    ' on ground
    msg$(0, 0) = "NOSE GEAR DNLMT"
    msg$(0, 0) = "SMALL GAP .04 G"

    msg$(0, 0) = "FAULTS FOR "
    msg$(0, 0) = "LEG NUMBER 42?"
    ' (yes/no, etc)

    ' Fault #11 = Leg 42 Left Main Gear Up Limit Too Far in air,
    ' Rigging Required
    msg$(0, 0) = "LEFT GEAR UPLMT"
    msg$(0, 0) = "LARGE GAP .18 A"

END SUB

SUB PresentFaults :
    DIM msg$(2, 8)
    DIM inchar$
    DIM mchoice
    DIM mchoices
   
    ' Fault Count = 3
    msg$(0, 0) = "PRESENT FAULTS:"
    msg$(1, 0) = "03"

    ' Fault #1 = Nose Gear Up Limit Sensor Open Circuit on ground
    msg$(0, 1) = "NOSE GEAR UPLMT"
    msg$(1, 1) = "SENSOR OPEN "

    ' Fault #2 = Nose Gear Down Limit Sensor Rigged Too Close (gap .05)
    ' on ground
    msg$(0, 2) = "NOSE GEAR DNLMT"
    msg$(1, 2) = "SMALL GAP .04 "

    ' Fault #3 = Left Main Gear Down Lock Conditioner Offset (code F1)
    ' requiring board A1 to be replaced
    msg$(0, 3) = "LEFT GEAR DNLMT"
    msg$(1, 3) = "BOARD A1 F1 "
   
    mchoices = 4

    inchar$ = "n"
    mchoice = 0
    WHILE inchar$ <> "q"
        SELECT CASE inchar$
            CASE IS = "u"
                mchoice = (mchoice + mchoices - 1) MOD (mchoices)
            CASE IS = "d"
                mchoice = (mchoice + 1) MOD (mchoices)
            CASE IS = "m"
                EXIT SUB
        END SELECT
        CLS
        LOCATE 20, 32
        PRINT msg$(0, mchoice)
        LOCATE 21, 32
        PRINT msg$(1, mchoice)
        DO
           inchar$ = INKEY$
        LOOP WHILE inchar$ = ""
    WEND
    CLS
END SUB

SUB PresentStatus
    DIM msg$(2, 20)
    DIM inchar$
    DIM mchoice
    DIM mchoices
   
    msg$(0, 0) = "RIGHT GEAR DNLMT"
    msg$(1, 0) = "DOWN & LOCKED"

    msg$(0, 1) = "RIGHT GEAR DNLMT"
    msg$(1, 1) = "TARGET NEAR"

    msg$(0, 2) = "RIGHT GEAR DNLMT"
    msg$(1, 2) = ".05 .12(0.10).20 "

    msg$(0, 3) = "RIGHT GEAR UPLMT"
    msg$(1, 3) = "NOT UP"

    msg$(0, 4) = "RIGHT GEAR UPLMT"
    msg$(1, 4) = "TARGET FAR"

    msg$(0, 5) = "RIGHT GEAR UPLMT "
    msg$(1, 5) = ".05 .12(>.30).20"

    msg$(0, 6) = "LEFT GEAR DNLMT"
    msg$(1, 6) = "DOWN & LOCKED"

    msg$(0, 7) = "LEFT GEAR DNLMT"
    msg$(1, 7) = "LARGE GAP "

    msg$(0, 8) = "LEFT GEAR DNLMT"
    msg$(1, 8) = ".05 .12(0.25).20"

    msg$(0, 9) = "LEFT GEAR UPLMT"
    msg$(1, 9) = "NOT UP"

    msg$(0, 10) = "LEFT GEAR UPLMT"
    msg$(1, 10) = "TARGET FAR"

    msg$(0, 11) = "LEFT GEAR UPLMT"
    msg$(1, 11) = ".05 .12(>.30).20"

    msg$(0, 12) = "NOSE GEAR"
    msg$(1, 12) = "DOWN & LOCKED"

    msg$(0, 13) = "NOSE GEAR DNLMT"
    msg$(1, 13) = "SMALL GAP "

    msg$(0, 14) = "NOSE GEAR DNLMT"
    msg$(1, 14) = ".05 .12(0.04).20"

    msg$(0, 15) = "NOSE GEAR UPLMT"
    msg$(1, 15) = "STATE UNKNOWN"

    msg$(0, 16) = "NOSE GEAR UPLMT"
    msg$(1, 16) = "SENSOR FAILOPEN"

    mchoices = 17

    inchar$ = "n"
    mchoice = 0
    WHILE inchar$ <> "q"
        SELECT CASE inchar$
            CASE IS = "u"
                mchoice = (mchoice + mchoices - 1) MOD (mchoices)
            CASE IS = "d"
                mchoice = (mchoice + 1) MOD (mchoices)
            CASE IS = "m"
                EXIT SUB
        END SELECT
        CLS
        LOCATE 20, 32
        PRINT msg$(0, mchoice)
        LOCATE 21, 32
        PRINT msg$(1, mchoice)
        DO
           inchar$ = INKEY$
        LOOP WHILE inchar$ = ""
    WEND
END SUB

SUB SelfTest :
    CLS
    PRINT "Inside E"

END SUB

SUB SystemOutputs :
    DIM msg$(2, 8)
    DIM inchar$
    DIM mchoice
    DIM mchoices
   
    msg$(0, 0) = "RIGHT GEAR"
    msg$(1, 0) = "SAFE"
    msg$(0, 1) = "LEFT GEAR"
    msg$(1, 1) = "SAFE"
    msg$(0, 2) = "NOSE GEAR"
    msg$(1, 2) = "UNSAFE"

    mchoices = 3

    inchar$ = "n"
    mchoice = 0
    WHILE inchar$ <> "q"
        SELECT CASE inchar$
            CASE IS = "u"
                mchoice = (mchoice + mchoices - 1) MOD (mchoices)
            CASE IS = "d"
                mchoice = (mchoice + 1) MOD (mchoices)
            CASE IS = "m"
                EXIT SUB
        END SELECT
        CLS
        LOCATE 20, 32
        PRINT msg$(0, mchoice)
        LOCATE 21, 32
        PRINT msg$(1, mchoice)
        DO
           inchar$ = INKEY$
        LOOP WHILE inchar$ = ""
    WEND
END SUB

