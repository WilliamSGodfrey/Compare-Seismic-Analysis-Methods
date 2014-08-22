Attribute VB_Name = "CTI_Info_PCACol_V4_1"
'pcaColumn Version 4.1
'
'pcaColumn Text Input (CTI) file format
'
'pcaColumn is able to read and save input data file into two formats, COL file or
'CTI file. CTI files are plain text files that can be edited by any text editing
'software.
'
'Caution must be used when editing a CTI file because some values may be
'interrelated. If one of these values is changed, then other interrelated values should
'be changed accordingly. While this is done automatically when a model is edited
'in the pcaColumn user graphic user interface (GUI), one must update all the
'related values in a CTI file manually in order to obtain correct results. For
'example, if units are changed from English to Metric in GUI, all the related input
'values are updated automatically. If this is done by editing a CTI file, however, not
'only the unit flag but also all the related input values must be updated manually.
'
'The best way to create a CTI file is by using the pcaColumn GUI and selecting
'CTI file type in the Save As menu command. Then, any necessary modifications to
'the CTI file can be applied with any text editor. However, it is recommended that
'users always verify modified CTI files by loading them in the pcaColumn GUI to
'ensure that the modifications are correct before running manually revised CTI files
'in batch mode.
'
'A CTI file is organized by sections. Each section contains a title in square
'brackets, followed by values required by the section. A CTI file contains the
'following sections.
'
'[pcaColumn Version]
'[Project]
'[Column ID]
'[Engineer]
'[Investigation Run Flag]
'[Design Run Flag]
'[Slenderness Flag]
'[User Options]
'[Irregular Options]
'[Ties]
'[Investigation Reinforcement]
'[Design Reinforcement]
'[Investigation Section Dimensions]
'[Design Section Dimensions]
'[Material Properties]
'[Reduction Factors]
'[Design Criteria]
'[External Points]
'[Internal Points]
'[Reinforcement Bars]
'[Factored Loads]
'[Slenderness: Column]
'[Slenderness: Column Above And Below]
'[Slenderness: Beams]
'[EI]
'[SldOptFact]
'[Phi_Delta]
'[Cracked I]
'[Service Loads]
'[Load Combinations]
'[BarGroupType]
'[User Defined Bars]
'
'Each section of a CTI file and allowable values of each parameter are described in
'details below. Corresponding GUI commands are presented in parenthesis.
'#pcaColumn Text Input (CTI) File
'
'A number sign, #, at the beginning of a line of text indicates that the line of text is
'comment. The # sign must be located at the beginning of a line. Comments may be
'added anywhere necessary in a CTI file to make the file more readable. If a
'comment appears in multiple lines, each line must be started with a # sign.
'
'[pcaColumn Version]
'    Reserved. Do not edit.
'
'[Project]
'    There is one line of text in this section.
'    Project name (menu Input | General Information…)
'
'[Column ID]
'    There is one line of text in this section.
'    Column ID (menu Input | General Information…)
'
'[Engineer]
'    There is one line of text in this section.
'    Engineer name (Menu Input | General Information…)
'
'[Investigation Run Flag]
'    Reserved. Do not edit.
'
'[Design Run Flag]
'    Reserved. Do not edit.
'
'[Slenderness Flag]
'    Reserved. Do not edit.
'
'[User Options]
'There are 26 values separated by commas in one line in this section. These values
'are described below in the order they appear from left to right.
'
'    1. 0-Investigation Mode; 1-Design Mode; (Run Option on menu Input |
'        General Information…)
'    2. 0-English Unit; 1-Metric Units; (Units on menu Input | General
'        Information…)
'    3. 0-ACI 318-02; 1- CSA A23.3-94; 2-ACI 318-05; 3-CSA A23.3-04; (Design
'        Code on menu Input | General Information…)
'    4. 0-X Axis Run; 1-Y Axis Run; 2-Biaxial Run; (Run Axis on menu Input |
'        General Information…)
'    5. Reserved. Do not edit.
'    6. 0-Slenderness is not considered; 1-Slenderness in considered; (Consider
'        slenderness? on menu Input | General Information…)
'    7. 0-Design for minimum number of bars; 1-Design for minimum area of
'        reinforcement; (Bar selection on menu Input | Reinforcement | Design
'        Criteria…)
'    8. Reserved. Do not edit.
'    9. 0-Rectangular Column Section; 1-Circular Column Section; 2-Irregular
'        Column Section; (menu Input | Section)
'    10. 0-Rectangular reinforcing bar layout; 1-Circular reinforcing bar layout; (Bar
'        Layout on menu Input | Reinforcement | All Sides Equal)
'    11. 0-Structural Column Section; 1-Architectural Column Section; 2-User
'        Defined Column Section; (Column Type on menu Input | Reinforcement |
'        Design Criteria…)
'    12. 0-Tied Confinement; 1-Spiral Confinement; 2-Other Confinement;
'        (Confinement drop-down list on menu Input | Reinforcement |
'        Confinement…)
'    13. Load type for investigation mode: (menu Input | Loads)
'        0-Factored; 1-Service; 2-Control Points; 3-Axial Loads
'    14. Load type for design mode: (menu Input | Loads)
'        0-Factored; 1-Service; 2-Control Points; 3-Axial Loads
'    15. Reinforcement layout for investigation mode: (menu Input | Reinforcement)
'        0-All Side Equal; 1-Equal Spacing; 2-Sides Different; 3-Irregular Pattern
'    16. Reinforcement layout for design mode: (menu Input | Reinforcement)
'        0-All Side Equal; 1-Equal Spacing; 2-Sides Different; 3-Irregular Pattern
'    17. Reserved. Do not edit.
'    18. Number of factored loads (menu Input | Loads | Factored…)
'    19. Number of service loads (menu Input | Loads | Service…)
'    20. Number of points on exterior column section
'    21. Number of points on interior section opening
'    22. Reserved. Do not edit.
'    23. Reserved. Do not edit.
'    24. Cover type for investigation mode: (menu Input | Reinforcement)
'        0-To transverse bar; 1-To longitudinal bar
'    25. Cover type for design mode: (menu Input | Reinforcement)
'        0-To transverse bar; 1-To longitudinal bar
'    26. Number of load combinations; (menu Input | Load | Load Combinations…)
'
'[Irregular Options]
'There are 13 values separated by commas in one line in this section. These values
'are described below in the order they appear from left to right. (menu Input |
'Section | Irregular | Section Editor menu Main | Drawing Area)
'    1. Reserved. Do not edit.
'    2. Reserved. Do not edit.
'    3. Reserved. Do not edit.
'    4. Reserved. Do not edit.
'    5. Area of reinforcing bar that is to be added through irregular section editor
'    6. Maximum X value of drawing area of irregular section editor
'    7. Maximum Y value of drawing area of irregular section editor
'    8. Minimum X value of drawing area of irregular section editor
'    9. Minimum Y value of drawing area of irregular section editor
'    10. Grid step in X of irregular section editor
'    11. Grid step in Y of irregular section editor
'    12. Grid snap step in X of irregular section editor
'    13. Grid snap step in Y of irregular section editor
'
'[Ties]
'There are 3 values separated by commas in one line in this section. These values
'are described below in the order they appear from left to right. (Menu Input |
'Reinforcement | Confinement…)
'    1. Index (0 based) of tie bars for longitudinal bars smaller that the one
'        specified in the 3rd item in this section in the drop-down list
'    2. Index (0 based) of tie bars for longitudinal bars bigger that the one specified
'        in the 3rd item in this section in the drop-down list
'    3. Index (0 based) of longitudinal bar in the drop-down list
'
'[Investigation Reinforcement]
'This section applies to investigation mode only. There are 12 values separated by
'commas in one line in this section. These values are described below in the order
'they appear from left to right.
'
'If Side Different (Menu Input | Reinforcement | Side Different…) is selected:
'    1. Number of top bars
'    2. Number of bottom bars
'    3. Number of left bars
'    4. Number of right bars
'    5. Index (0 based) of top bars (Top Bar Size drop-download list)
'    6. Index (0 based) of bottom bars (Bottom Bar Size drop-download list)
'    7. Index (0 based) of left bars (Left Bar Size drop-download list)
'    8. Index (0 based) of right bars (Right Bar Size drop-download list)
'    9. Clear cover to top bars
'    10. Clear cover to bottom bars
'    11. Clear cover to left bars
'    12. Clear cover to right bars
'
'If All Sides Equal (Menu Input | Reinforcement | All Sides Equal…) or Equal
'Spacing (Menu Input | Reinforcement | Equal Spacing…) is selected:
'    1. Number of bars (No. of Bars text box)
'    2. Reserved. Do not edit.
'    3. Reserved. Do not edit.
'    4. Reserved. Do not edit.
'    5. Index (0 based) of bar (Bar Size drop-down list)
'    6. Reserved. Do not edit.
'    7. Reserved. Do not edit.
'    8. Reserved. Do not edit.
'    9. Clear cover to bar (Clear Cover text box)
'    10. Reserved. Do not edit.
'    11. Reserved. Do not edit.
'    12. Reserved. Do not edit.
'
'If Irregular Pattern (Menu Input | Reinforcement | Irregular Pattern…) is selected:
'    Reserved. Do not edit.
'
'[Design Reinforcement]
'This section applies to design mode only. There are 12 values separated by
'commas in one line in this section. These values are described below in the order
'they appear from left to right.
'
'If Side Different (Menu Input | Reinforcement | Side Different…) is selected:
'    1. Minimum number of top and bottom bars
'    2. Maximum number of top and bottom bars
'    3. Minimum number of left and right bars
'    4. Maximum number of left and right bars
'    5. Index (0 based) of minimum size for top and bottom bars
'    6. Index (0 based) of maximum size for top and bottom bars
'    7. Index (0 based) of minimum size for left and right bars
'    8. Index (0 based) of maximum size for left and right bars
'    9. Clear cover to top and bottom bars
'    10. Reserved. Do not edit.
'    11. Clear cover to left and right bars
'    12. Reserved. Do not edit.
'
'If All Sides Equal (Menu Input | Reinforcement | All Sides Equal…) or Equal
'Spacing (Menu Input | Reinforcement | Equal Spacing…) is selected:
'    1. Minimum number of bars
'    2. maximum number of bars
'    3. Reserved. Do not edit.
'    4. Reserved. Do not edit.
'    5. Index (0 based) of minimum size of bars
'    6. Index (0 based) of maximum size of bars
'    7. Reserved. Do not edit.
'    8. Reserved. Do not edit.
'    9. Clear cover
'    10. Reserved. Do not edit.
'    11. Reserved. Do not edit.
'    12. Reserved. Do not edit.
'
'[Investigation Section Dimensions]
'This section applies to investigation mode only. There are 2 values separated by
'commas in one line in this section. These values are described below in the order
'they appear from left to right.
'
'If rectangular section (Menu Input | Section | Rectangular…) is selected:
'    1. Section width (along X)
'    2. Section depth (along Y)
'
'If circular section (Menu Input | Section | Circular…) is selected:
'    1. Section diameter
'    2. Reserved. Do not edit.
'
'If irregular section (Menu Input | Section | Irregular) is selected:
'    1. Reserved. Do not edit.
'    2. Reserved. Do not edit.
'
'[Design Section Dimensions]
'This section applies to design mode only. There are 6 values separated by commas
'in one line in this section. These values are described below in the order they
'appear from left to right.
'
'If rectangular section (Menu Input | Section | Rectangular…) is selected:
'    1. Section width (along X) Start
'    2. Section depth (along Y) Start
'    3. Section width (along X) End
'    4. Section depth (along Y) End
'    5. Section width (along X) Increment
'    6. Section depth (along Y) Increment
'    If circular section (Menu Input | Section | Circular…) is selected:
'    1. Diameter start
'    2. Reserved. Do not change.
'    3. Diameter end
'    4. Reserved. Do not change.
'    5. Diameter increment
'    6. Reserved. Do not change.
'
'[Material Properties]
'There are 8 values separated by commas in one line in this section. These values
'are described below in the order they appear from left to right. (Menu Input |
'Material Properties…)
'
'    1. Concrete strength, f’c
'    2. Concrete modulus of elasticity, Ec
'    3. Concrete maximum stress, fc
'    4. Beta(1) for concrete stress block
'    5. Concrete ultimate strain
'    6. Steel yield strength, fy
'    7. Steel modulus of elasticity, Es
'    8. Reserved. Do not edit.
'
'[Reduction Factors]
'There are 4 values separated by commas in one line in this section. These values
'are described below in the order they appear from left to right. (Menu Input |
'Reinforcement | Confinement…)
'
'    1. Phi(a) for axial compression
'    2. Phi(b) for tension-controlled failure
'    3. Phi(c) for compression-controlled failure
'    4. Reserved. Do not edit.
'
'[Design Criteria]
'There are 4 values separated by commas in one line in this section. These values
'are described below in the order they appear from left to right. (Menu Input |
'Reinforcement | Design Criteria…)
'
'    1. Minimum reinforcement ratio
'    2. Maximum reinforcement ratio
'    3. Minimum clear spacing between bars
'    4. Design/Required ratio
'
'[External Points]
'This section applies to irregular section in investigation mode only. The first line
'contains the number of points on exterior section perimeter. Each of the following
'lines contains 2 values of X and Y coordinates (separated by comma) of a point.
'
'    Number of Points, n
'    Point_1_X , Point_1_Y
'    Point_2_X , Point_2_Y
'    .
'    .
'    .
'    Point_n_X , Point_n_Y
'
'[Internal Points]
'This section applies to irregular section in investigation mode only. The first line
'contains the number of points on an interior opening perimeter. Each of the
'following lines contains 2 values of X and Y coordinates (separated by comma) of
'a point. If no opening exists, then the first line must be 0.
'
'    Number of Points, n
'    Point_1_X , Point_1_Y
'    Point_2_X , Point_2_Y
'    .
'    .
'    .
'    Point_n_X , Point_n_Y
'
'[Reinforcement Bars]
'This section applies to irregular section in investigation mode only. The first line
'contains the number of reinforcing bars. Each of the following lines contains 3
'values of area, X and Y coordinates (separated by comma) of a bar.
'
'    Number of bars, n
'    Bar_1_area , Bar_1_X, Bar_1_Y
'    Bar_2_area , Bar_2_X, Bar_2_Y
'    .
'    .
'    .
'    Bar_n_area , Bar_n_X, Bar_n_Y
'
'[Factored Loads]
'The first line contains the number of factored loads defined. Each of the following
'lines contains 3 values of axial load, X-moment, and Y-moment separated by
'commas. (Menu Input | Loads | Factored Loads…)
'
'    Number of Factored Loads, n
'    Load_1 , x - Moment_1, y - Moment_1
'    Load_2 , x - Moment_2, y - Moment_2
'    .
'    .
'    .
'    Load_n , x - Moment_n, y - Moment_n
'
'[Slenderness: Column]
'This section contains 2 lines describing slenderness parameters for column being
'designed. The first line is for X-axis parameters, and the second line is for Y-axis
'parameters.
'There are 5 values separated by commas in each line. These values are described
'below in the order they appear from left to right. (Menu Input | Slenderness |
'Design Column…)
'    1. Column clear height
'    2. k(braced)
'    3. k(sway)
'    4. 0-Non-sway frame; 1-Sway frame
'    5. 0-Compute ‘k’ factors; 1-Input k factors
'
'[Slenderness: Column Above And Below]
'This section contains 2 lines describing slenderness parameters for column above
'and column below. The first line is for column above, and the second line is for
'column below. (Menu Input | Slenderness | Columns Above/Below…)
'
'here are 6 values separated by commas in line 1 for column above. These values
'are described below in the order they appear from left to right.
'    1. 0-Column specified; 1-No column specified
'    2. Column Height
'    3. Column width (along X)
'    4. Column depth (along Y)
'    5. Concrete compressive strength, f’c
'    6. Concrete modulus of elasticity, Ec
'
'There are 6 values separated by commas in line 2 for column below. These values
'are described below in the order they appear from left to right.
'    1. 0-Column specified; 1-No column specified
'    2. Column Height
'    3. Column width (along X)
'    4. Column depth (along Y)
'    5. Concrete compressive strength, f’c
'    6. Concrete modulus of elasticity, Ec
'
'[Slenderness: Beams]
'This section contains 8 lines. Each line describes a beam.
'
'    Line 1: X-Beam (perpendicular to X), Above Left
'    Line 2: X-Beam (perpendicular to X), Above Right
'    Line 3: X-Beam (perpendicular to X), Below Left
'    Line 4: X-Beam (perpendicular to X), Below Right
'    Line 5: Y-Beam (perpendicular to Y), Above Left
'    Line 6: Y-Beam (perpendicular to Y), Above Right
'    Line 7: Y-Beam (perpendicular to Y), Below Left
'    Line 8: Y-Beam (perpendicular to Y), Below Right'
'
'        There are 7 values separated by commas for each beam in each line. (Menu Input |
'        Slenderness | X-Beams…, Input | Slenderness | Y-Beams…) These values are
'        described below in the order they appear from left to right.
'
'        1. 0-beam specified; 1-no beam specified
'        2. Beam span length (c/c)
'        3. Beam width
'        4. Beam depth
'        5. Beam section moment of inertia
'        6. Concrete compressive strength, f’c
'        7. Concrete modulus of elasticity, Ec
'
'[EI]
'    Reserved. Do not edit.
'
'[SldOptFact]
'There is 1 value in this section for slenderness factors. (Code Default and User
'Defined radio buttons on menu Input | Slenderness | factors…)
'
'    0-Code default; 1-User defined
'
'[Phi_Delta]
'There is 1 value in this section for slenderness factors. (Menu Input | Slenderness |
'factors…)
'
'    Stiffness reduction factor
'
'[Cracked I]
'There are 2 values separated by commas in one line in this section. These values
'are described below in the order they appear from left to right. (Menu Input |
'Slenderness | factors…)
'
'    1. Beam cracked section coefficient
'    2. Column cracked section coefficient
'
'[Service Loads]
'This section describes defined service loads. (Menu Input | Loads | Service…) The
'first line contains the number of service loads. Each of the following lines contains
'values for one service load.
'There are 25 values for each service load in one Line separated by commas. These
'values are described below in the order they appear from left to right.
'    1. Dead Axial Load
'    2. Dead X-moment at top
'    3. Dead X-moment at bottom
'    4. Dead Y-moment at top
'    5. Dead Y-moment at bottom
'    6. Live Axial Load
'    7. Live X-moment at top
'    8. Live X-moment at bottom
'    9. Live Y-moment at top
'    10. Live Y-moment at bottom
'    11. Wind Axial Load
'    12. Wind X-moment at top
'    13. Wind X-moment at bottom
'    14. Wind Y-moment at top
'    15. Wind Y-moment at bottom
'    16. EQ. Axial Load
'    17. EQ. X-moment at top
'    18. EQ. X-moment at bottom
'    19. EQ. Y-moment at top
'    20. EQ. Y-moment at bottom
'    21. Snow Axial Load
'    22. Snow X-moment at top
'    23. Snow X-moment at bottom
'    24. Snow Y-moment at top
'    25. Snow Y-moment at bottom
'
'[Load Combinations]
'This section describes defined load combinations. (Menu Input | Loads | Load
'Combinations…) The first line contains the number of load combinations. Each of
'the following lines contains load factors for one load combination.
'
'    Number of load combinations, n
'    Dead_1, Live_1, Wind_1, E.Q._1, Snow_1
'    Dead_2, Live_2, Wind_2, E.Q._2, Snow_2
'    .
'    .
'    .
'    Dead_n, Live_n, Wind_n, E.Q._n, Snow_n
'
'[BarGroupType]
'There is 1 value in this section. (Bar Set drop-down list on menu Options |
'Reinforcement…)
'
'    0-User defined
'    1-ASTM615
'    2-CSA G30.18
'    3-prEN 10080
'    4-ASTM615M
'
'[User Defined Bars]
'This section contains user-defined reinforcing bars. (Menu Options |
'Reinforcement…) The first line contains the number of defined bars. Each of the
'following lines contains values for one bar separated by commas.
'
'    Number of user-defined bars, n
'    Bar_1_size , Bar_1_diameter, Bar_1_area, Bar_1_weight
'    Bar_2_size , Bar_2_diameter, Bar_2_area, Bar_2_weight
'    .
'    .
'    .
'    Bar_n_size , Bar_n_diameter, Bar_n_area, Bar_n_weight

