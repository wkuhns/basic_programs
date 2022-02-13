{--------------------------------------------------------------------------

FILE: robot1.pas

This file contains the main program for level 1 of the robot project. It also
contains solutions for the following level 1 tasks:

	Calibratio file I/O
	Calibrate analog sensors
	Convert sensor values to degrees
	
---------------------------------------------------------------------------}

program robot;

uses dos, crt, graph, gamecard, parallel;


{
     CalArray is a global array of sensor calibration values. The first
     index is the channel number (1 through 4). There is a 2x5 array
     that contains angle vs. sensor reading values for each channel.
}

var
    CalArray : array[1..4, 1..2, 1..5] of real;
    DirectionArray : array[1..4, 1..2] of integer;

{--------------------- procedures and functions --------------------------}

{---------------------------------------------------------------------------

  Solution to Convert analog sensor values to degrees

  This solution uses linear interpolation.

  By definition, the angle values in column 1 of CalArray are ascending, but
  the sensor readings in column 2 may be either ascending or descending.

  This algorithm searches through CalArray until it finds two angle/sensor
  value pairs with sensor values above and below the value passed to this 
  function. These pairs are stored in a1/s1 and a2/s2.

  It then calculates an angle value via standard linear interpolation.

---------------------------------------------------------------------------}

function GetDegrees(channel, sensor : integer) : real;

var
    s1, s2, a1, a2 : real;              { table value storage }
    i : integer;

begin
    if CalArray[channel,2,1] > CalArray[channel,2,2] then
    begin                       { descending table }
        i := 5;
        while (CalArray[channel,2,i] < sensor) and (i > 1) do i := i - 1;
        a1 := CalArray[channel,1,i];
        s1 := CalArray[channel,2,i];
        a2 := CalArray[channel,1,i+1];
        s2 := CalArray[channel,2,i+1];
        GetDegrees := (((s1-sensor) * (a2-a1)) / (s1-s2)) + a1;
    end;

    if CalArray[channel,2,1] < CalArray[channel,2,2] then
    begin
        i := 1;
        while (CalArray[channel,2,i] < sensor) and (i < 5) do i := i + 1;
        a1 := CalArray[channel,1,i-1];
        s1 := CalArray[channel,2,i-1];
        a2 := CalArray[channel,1,i];
        s2 := CalArray[channel,2,i];
        GetDegrees := (((sensor-s1) * (a2-a1)) / (s2-s1)) + a1;
    end;
end;

{---------------------------------------------------------------------------

  This procedure is used by the main program to update sensor and angle
  values on the screen.

---------------------------------------------------------------------------}

procedure ShowStats;

var
   i, sensor : integer;
   angle : real;
begin
     for i := 0 to 2 do
     begin
          sensor := ReadAnalogSensor(i+1);
          angle := GetDegrees(i+1, sensor);
          gotoxy(20 * i + 11, 2);
          writeln(sensor:3);
          gotoxy(20 * i + 8, 4);
          writeln(angle:6:1);
    end;
end;

{---------------------------------------------------------------------------

  Solution to automatic joint position task.

  Uses global DirectionArray to determine which direction to drive motor
  to reach commanded joint angle.

  This version slows the motor as it approaches the target angle.

---------------------------------------------------------------------------}

procedure SetPosition(channel : integer; angle : real);

var
    currentangle : real;		{ where we are right now }
    delta : real;			{ how far we have to go }
    direction : integer;		{ the motor direction we need }
    power : integer;
    decrease : boolean;

const
    range = 20;			{ slow down if less than 20 degrees }

begin
    currentangle := GetDegrees(channel,ReadAnalogSensor(channel));

    delta := abs(currentangle-angle);

    decrease := currentangle > angle;

    if decrease then			{ we are decreasing the angle }
    begin
        direction := DirectionArray[channel,2];
        while currentangle > angle do
        begin
            delta := currentangle - angle;
            power := trunc(delta / range * 7 + 3);
            if (delta > range) then power := 10;
            if (power = 0) then power := 1;
            SetMotor(channel,direction,1);
            delay(power);                           { motor on time }
            SetMotor(channel,direction,0);
            delay(10-power);                        { motor off time }
            currentangle := GetDegrees(channel,ReadAnalogSensor(channel));
            ShowStats;				    { keep user informed }
        end;
        SetMotor(channel,direction,0);
    end;

    if not decrease then		{ we are increasing the angle }
    begin
        direction := DirectionArray[channel,1];
        while currentangle < angle do
        begin
            delta := angle - currentangle;
            power := trunc(delta / range * 7 + 3);
            if (delta > range) then power := 10;
            if (power = 0) then power := 1;
            SetMotor(channel,direction,1);
            delay(power);                           { motor on time }
            SetMotor(channel,direction,0);
            delay(10-power);                        { motor off time }
            currentangle := GetDegrees(channel,ReadAnalogSensor(channel));
            ShowStats;
        end;
        SetMotor(channel,direction,0);
    end;
end;

{---------------------------------------------------------------------------

  Solution to manual control exercise.

  This solution incorporates options to increase or decrease to motor speed.

---------------------------------------------------------------------------}

procedure ManualControl(joint : integer);
var
    quitflag : char;
    direction, speed : integer;

begin
    quitflag := 'x';             { never this value in the loop }
    direction := 0;
    speed := 0;

    gotoxy(10,17);
    Writeln('Manually controlling channel ',joint);
    gotoxy(10,19);
    writeln('Quit, Faster, Slower, Reverse, Pause');

    while quitflag <> 'q' do
    begin
        SetMotor(joint,direction,1);
        delay(speed);                           { motor on time }
        SetMotor(joint,direction,0);
        delay(10-speed);                        { motor off time }
        ShowStats;
        if KeyPressed then
        begin
            quitflag := ReadKey;
            case quitflag of
                'f': speed := speed + 1;
                's': speed := speed - 1;
                'r': direction := direction + 1;
                'p': speed := 0;
            end;
            if speed > 10 then speed := 10;
            if speed < 0 then speed := 0;
            if direction = 2 then direction := 0;
        end;
    end;
    SetMotor(joint,direction,0);
end;

{---------------------------------------------------------------------------

  Solution to get calibration task.

  User is prompted to move the joint to each angle in CalArray.
  ManualControl is used to allow the user to position the joint.
  When ManualControl is finished, the joint position is read into CalArray.

---------------------------------------------------------------------------}

procedure GetCalibration(channel : integer);

var
    i : integer;

begin

    for i := 1 to 5 do
    begin
        gotoxy(10,15);
        write('Move joint ', channel, ' to ');
	write(CalArray[channel,1,i]:5:1, ' degrees');
        ManualControl(channel);
        CalArray[channel,2,i] := ReadAnalogSensor(channel);
    end;
end;

{---------------------------------------------------------------------------

  Solutions to Calibration file I/O task.

---------------------------------------------------------------------------}

procedure ReadCalFromDisk(channel : integer);

var
    infile : Text;
    i : integer;

begin
    case channel of
	1: assign(infile, 'cal1.dat');
	2: assign(infile, 'cal2.dat');
	3: assign(infile, 'cal3.dat');
	4: assign(infile, 'cal4.dat');
    end;

    FileMode := 0;                      {read only}
    reset(infile);

    for i := 1 to 5 do
    begin
	readln(infile, CalArray[channel,1,i], CalArray[channel,2,i]);
    end;

    close(infile);
end;

{---------------------------------------------------------------------------}

procedure WriteCalToDisk(channel : integer);

var
    outfile : Text;
    i : integer;

begin
    case channel of
	1: assign(outfile, 'cal1.dat');
	2: assign(outfile, 'cal2.dat');
	3: assign(outfile, 'cal3.dat');
	4: assign(outfile, 'cal4.dat');
    end;

    FileMode := 1;                      {write only}
    rewrite(outfile);

    for i := 1 to 5 do
    begin
	writeln(outfile, CalArray[channel,1,i]:5:1,CalArray[channel,2,i]:6:1);
    end;

    close(outfile);
end;

{---------------------------------------------------------------------------

  Used by main program. Prints portions of the screen that don't change.

---------------------------------------------------------------------------}

procedure paintmenu;

var
   i : integer;

begin
     for i := 0 to 3 do
     begin
          gotoxy(20 * i + 2, 2);
          writeln('Sensor');
          gotoxy(20 * i + 2, 4);
          writeln('Angle');
     end;
     gotoxy(10, 8);
     writeln('1: Select channel 1');
     gotoxy(10, 9);
     writeln('2: Select channel 2');
     gotoxy(10, 10);
     writeln('3: Select channel 3');
     gotoxy(10,11);
     writeln('s: Set position for channel');
     gotoxy(50, 8);
     writeln('w : Write cal data to disk');
     gotoxy(50, 9);
     writeln('r : Recalibrate channel');
     gotoxy(50, 10);
     writeln('m : Manually move channel');

     gotoxy(50, 11);
     writeln('q : Quit');
end;

{----------------------- body of robot main program ----------------------}

var
    i, choice, channel : integer;
    angle : real;
    quitflag : char;
    GraphDriver, GraphMode, ErrorCode : integer;

begin

    GraphDriver := Detect;
    InitGraph(GraphDriver, GraphMode, 'E:\TP\BGI');
    ErrorCode := GraphResult;
    if ErrorCode <> grOk then
    begin
        writeln('Graphics error: ', GraphErrorMsg(ErrorCode));
        halt(1);
    end;
    CloseGraph;

    for i := 1 to 4 do
    begin
        ReadCalFromDisk(i);
        SetMotor(i,0,0);
    end;

    { ---- These values will depend on how the motors are wired. ---- }

    DirectionArray[1,1] := 1;
    DirectionArray[1,2] := 0;
    DirectionArray[2,1] := 1;
    DirectionArray[2,2] := 0;
    DirectionArray[3,1] := 1;
    DirectionArray[3,2] := 0;

    channel := 1;
{    cleardevice;  }
    clrscr;
    window(1,1,80,24);
    paintmenu;
    while quitflag <> 'q' do
    begin
        ShowStats;
        if KeyPressed then
        begin
            quitflag := ReadKey;

            case quitflag of
                 '1': channel := 1;
                 '2': channel := 2;
                 '3': channel := 3;
                 's': begin
                         angle := 50;
                         while angle > 0 do
                         begin
                              gotoxy(10,15);
                              write('Enter angle, -1 to exit         ');
                              gotoxy(35,15);
                              read(angle);
                              if angle >= 0 then
                                 SetPosition(channel,angle);
                         end;
                    end;
                    'w': WriteCalToDisk(channel);            {write it back out}
                    'r': GetCalibration(channel);            {calibrate sensor}
                    'm': ManualControl(channel);
            end;  { end case }
            clrscr;
            paintmenu;
        end; {end keypressed }
    end; {while choice < 8 }
end.
