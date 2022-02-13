{--------------------------------------------------------------------------

FILE: gamecard.pas

This file contains example code for functions that access the game card used
in the robotics curriculum. It can be compiled using Borland Turbo Pascal to
create a .TPU file.

---------------------------------------------------------------------------}

unit GameCard;

interface

    function ReadAnalogSensor(channel : byte) : integer;
    function ReadGameSwitch(channel : byte) : boolean;
    procedure ReadJoysticks(var X1, Y1, X2, Y2 : word);

implementation

uses Dos;

var
    Reg : Registers;

{----------------------------- private routines ----------------------------}

{
  This function returns a byte in which the four most significant bits
  are the four switch positions
}

function ReadButtons : byte;
begin
    Reg.ah := $84;
    Reg.dx := 0;
    intr($15,Reg);
    ReadButtons := Reg.al;
end;

procedure OldReadJoysticks(var X1, Y1, X2, Y2 : word);
{ not working yet... }
begin
    asm
        MOV DX, 201H

        LES DI,X1
        MOV AH, 01
        CALL @1

        LES DI,Y1
        MOV AH,02
        CALL @1

        LES DI,X2
        MOV AH,04
        CALL @1

        LES DI,Y2
        MOV AH,08
        CALL @1

        JP @2
    @1:
        MOV CX, 0FFFFH
    @3:                         { WAIT TILL BIT IS 0 }
        IN AL,DX
        TEST AL,AH
        LOOPNE @3

        MOV CX, 0FFFFH
        CLI
        OUT DX,AL
    @4:
        IN AL,DX
        TEST AL, AH
        LOOPNE @4

        STI
        NEG CX
        ADD CX, 0FFFFH
        { ES }
        MOV [DI], CX
    @2:
    end;
end; { OldReadJoysticks }

{$F-}

{----------------------------- public routines -----------------------------}

{
  Solution to Robotics exercises.
}

{--------------------------------------------------------------------------
  Solution to 'Low level routine to read analog values'

  This procedure sets the value of four variables, one for each joystick
  axis.

  BIOS software interrupt vector 15(hex), function 84(hex) reads all four 
  joystick channels. The results are returned in register pairs ax, bx,
  cx, and dx.
---------------------------------------------------------------------------}

procedure ReadJoysticks(var X1, Y1, X2, Y2 : word);

begin
    Reg.ah := $84;
    Reg.dx := 1;
    intr($15,Reg);

    X1 := Reg.ax;
    Y1 := Reg.bx;
    X2 := Reg.cx;
    Y2 := Reg.dx;
end;

{---------------------------------------------------------------------------
  Solution to 'Low level routine to read game card switches'

  The switches are in the four high order bits of port 201(hex). The low
  bits are masked out with the 'and' operator.
---------------------------------------------------------------------------}

function ReadGameSwitches : byte;

begin
    ReadGameSwitches := (port[$201] and $F0);
end;

{---------------------------------------------------------------------------
  Solution to ReadSwitch in 'Embed low level routines' exercise.

  Depending on the channel desired, unwanted bits are masked out.
---------------------------------------------------------------------------}

function ReadGameSwitch(channel : byte) : boolean;

begin
    case channel of
        1: ReadGameSwitch := (ReadButtons and $10) = 0;
        2: ReadGameSwitch := (ReadButtons and $20) = 0;
        3: ReadGameSwitch := (ReadButtons and $40) = 0;
        4: ReadGameSwitch := (ReadButtons and $80) = 0;
    end;
end;

{---------------------------------------------------------------------------
  Solution to ReadAnalogSensor in 'Embed low level routines'.

  Use low level routine to get all four channels, then return the correct
  one to the caller.  
---------------------------------------------------------------------------}

function ReadAnalogSensor(channel : byte) : integer;

var
    CH1, CH2, CH3, CH4 : word;

begin
    ReadJoysticks(CH1, CH2, CH3, CH4);
    case channel of
        1: ReadAnalogSensor := CH1;
        2: ReadAnalogSensor := CH2;
        3: ReadAnalogSensor := CH3;
        4: ReadAnalogSensor := CH4;
    end;
end;

end.
