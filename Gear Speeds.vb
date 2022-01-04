Gear Speeds
---------------------------
=
IF( 
    A2=$V$13            1st = Input Speed * (18/52)
        B2*$AC$13      

IF(
    A2=$V$14            2nd = Input Speed * (22/46)
        B2*$AC$14

IF(
    A2=$V$15            3rd = Input Speed * (27/43)
        A2*$AC$15

IF(
    A2=$V$16            4th = Output Speed / (11/35) / (28/39) / (27/34)
        C2/$AC$16
        
IF(
    A2=$V$17            5th = Output Speed / (11/35) / (28/39) / (34/35)
        C2/$AC$17

IF(
    A2=$V$18            6th = Output Speed / (11/35) / (28/39) / (32/28)
        C2/$AC$18
        
IF(
    A2=$V$19            7th = Output Speed / (11/35) / (28/39) / (40/31)
        C2/$AC$19







Shaft Speeds
---------------------------

=
IF(
    OR(
        A2=$V$13    1st = Output Speed / (11/35) / (28/39)
        A2=$V$14    2nd = Output Speed / (11/35) / (28/39)
        A2=$V$15    3rd = Output Speed / (11/35) / (28/39)
    )
        C2/$AD$13   
  
    C2/$AD$16   = Input Shaft Speed
)



1st Gear
----------
GEAR_SPEED_2 = 9750 * RATIO_1
SHAFT_SPEED_2 = GEAR_SPEED_2 - DELTA
OUTPUT_SPEED_2 = SHAFT_SPEED_2 * (11/35) * (28/39)

OUTPUT_SPEED_2 = ((9750 * (18/52)) - DELTA) * (11/35) * (28/39)

2nd Gear
----------
GEAR_SPEED_2 = 9750 * RATIO_2
SHAFT_SPEED_2 = GEAR_SPEED_2 - DELTA
OUTPUT_SPEED_2 = SHAFT_SPEED_2 * (11/35) * (28/39)

OUTPUT_SPEED_2 = ((9750 * (22/46)) - DELTA) * (11/35) * (28/39)

3rd Gear
----------
GEAR_SPEED_2 = 9750 * RATIO_3
SHAFT_SPEED_2 = GEAR_SPEED_2 - DELTA
OUTPUT_SPEED_2 = SHAFT_SPEED_2 * (11/35) * (28/39)

OUTPUT_SPEED_2 = ((9750 * (27/43))-DELTA) * (11/35) * (28/39)

4th Gear
---------
SHAFT_SPEED_2 = 9750
GEAR_SPEED_2 = SHAFT_SPEED_2 - DELTA
OUTPUT_SPEED_2 = GEAR_SPEED_2 * RATIO_4 * (11/35) * (28/39)

OUTPUT_SPEED_2 = (9750 - DELTA) * (27/34) * (11/35) * (28/39)

INPUT - >
----------
=IF(
    GEAR=1,
        ((9750*(18/52))-DELTA)*(11/35)*(28/39),
IF(
    GEAR=2,
        ((9750*(22/46))-DELTA)*(11/35)*(28/39),
IF(
    GEAR=3,
        ((9750*(27/43))-DELTA)*(11/35)*(28/39),
IF(
    GEAR=4,
        (9750-DELTA)*(27/34)*(11/35)*(28/39),
IF(
    GEAR=5,
        (9750-DELTA)*(34/35)*(11/35)*(28/39),
IF(
    GEAR=6,
        (9750-DELTA)*(32/28)*(11/35)*(28/39),
IF(
    GEAR=7,
        (9750-DELTA)*(40/21)*(11/35)*(28/39),
)))))))


=IF(
    B2<9950,
        B2,
=IF(
    A2=1,
        ((9750*(18/52))-J2)*(11/35)*(28/39),
IF(
    A2=2,
        ((9750*(22/46))-J2)*(11/35)*(28/39),
IF(
    A2=3,
        ((9750*(27/43))-J2)*(11/35)*(28/39),
IF(
    A2=4,
        (9750-J2)*(27/34)*(11/35)*(28/39),
IF(
    A2=5,
        (9750-J2)*(34/35)*(11/35)*(28/39),
IF(
    A2=6,
        (9750-J2)*(32/28)*(11/35)*(28/39),
IF(
    A2=7,
        (9750-J2)*(40/21)*(11/35)*(28/39),
))))))))


OUTPUT SPEED
----------
=IF(
    GEAR=1, 
        (NEW GEAR SPEED - DELTA SPEED) / (11/35) * (28/39)      ' OUTPUT SPEED

)