program TEST;

uses gamecard;

var
    sensor_reading: integer;
    switch_reading: boolean;
    i: integer;

begin
     write(output, 'Sensor readings: ');
     for i := 1 to 4 do
     begin
         sensor_reading := ReadAnalogSensor(i);
         write(output, sensor_reading, '  ');
     end;
     writeln(output);

     write(output, 'Switch readings: ');
     for i := 1 to 4 do
     begin
         switch_reading := ReadGameSwitch(i);
         write(output, switch_reading, '  ');
     end;
     writeln(output);

end.