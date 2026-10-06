OptionExplicit

Public Enum Errors
    FE_OK = -1
    FE_FAIL = 0
    FE_CANCEL = 2
    FE_INVALID = 3
    FE_NOT_EXIST = 4
    FE_SECURITY = 5
    FE_NOT_AVAILABLE = 6
    FE_TOO_SMALL = 7
    FE_BAD_TYPE = 8
    FE_BAD_DATA = 9
    FE_NO_MEMORY = 10
    FE_NO_FILENAME = 16
End Enum

Public Enum Events
    FEVENT_INITIALIZE = 1
    FEVENT_NEWMODEL = 2
    FEVENT_ENDMODEL = 3
    FEVENT_SHUTDOWN = 4
    FEVENT_COMMAND = 5
End Enum

Public Enum OptionDefinitions
    FO_FILE_MODEL = 1
    FO_FILE_NEUTRAL = 2
    FO_FILE_ABAQUS = 3
    FO_FILE_ANSYS = 4
    FO_FILE_MSC_NASTRAN = 5
End Enum

Public Enum feAppMessage
    Normal = 0
    Highlight = 1
    Warning = 2
    Error = 3
End Enum

Public Enum entityTYPE
    FT_POINT = 3
    FT_CURVE = 4
    FT_SURFACE = 5
    FT_VOLUME = 6
    FT_NODE = 7
    FT_ELEM = 8
    FT_CSYS = 9
    FT_MATL = 10
    FT_PROP = 11
    FT_LOAD_DIR = 12
    FT_SURF_LOAD = 13
    FT_GEOM_LOAD = 14
    FT_NTHERM_LOAD = 15
    FT_ETHERM_LOAD = 16
    FT_BCO = 18
    FT_BCO_GEOM = 19
    FT_BEQ = 20
    FT_ESP_TEXT = 21
    FT_VIEW = 22
    FT_GROUP = 24
    FT_VAR = 27
    FT_OUT_CASE = 28
    FT_OUT_DIR = 29
    FT_OUT_DATA = 30
    FT_REPORT = 31
    FT_BOUNDARY = 32
    FT_LAYER = 33
    FT_MATL_TABLE = 34
    FT_FUNCTION_DIR = 35
    FT_FUNCTION_TABLE = 36
    FT_BC_DIR = 17
    FT_SOLID = 39
    FT_COLOR = 40
    FT_OUT_CSYS = 41
    FT_CONTACT = 58
    FT_GRTYPE = 59
    FT_AMGR_DIR = 60
    FT_TMG_BCO = 112
    FT_TMG_CONTROL = 113
    FT_TMG_INTEGER = 114
    FT_TMG_REAL = 115
    FT_TMG_OPTION = 116
End Enum

Public Enum GroupRules
    FGDEF_CSys_byDefCSys = 1
    FGDEF_CSys_byType = 2
    FGDEF_Point_ID = 3
    FGDEF_Point_byDefCSys = 4
    FGDEF_Point_onCurve = 5
    FGDEF_Curve_ID = 6
    FGDEF_Curve_byPoint = 7
    FGDEF_Curve_onSurface = 8
    FGDEF_Surface_ID = 9
    FGDEF_Surface_byCurve = 10
    FGDEF_Surface_onVolume = 11
    FGDEF_Volume_ID = 12
    FGDEF_Volume_bySurface = 13
    FGDEF_Text_ID = 14
    FGDEF_Boundary_ID = 15
    FGDEF_Boundary_byCurve = 16
    FGDEF_Node_ID = 17
    FGDEF_Node_byDefCSys = 18
    FGDEF_Node_byOutCSys = 19
    FGDEF_Node_onElem = 20
    FGDEF_Elem_ID = 21
    FGDEF_Elem_byMatl = 22
    FGDEF_Elem_byProp = 23
    FGDEF_Elem_byType = 24
    FGDEF_Elem_byNode = 25
    FGDEF_Matl_ID = 26
    FGDEF_Matl_onProp = 27
    FGDEF_Matl_onElem = 28
    FGDEF_Matl_byType = 29
    FGDEF_Prop_ID = 30
    FGDEF_Prop_onElem = 31
    FGDEF_Prop_byMatl = 32
    FGDEF_Prop_byType = 33
    FGDEF_Load_byNode = 34
    FGDEF_Load_byElem = 35
    FGDEF_BCo_ID = 36
    FGDEF_BEq_byNode = 37
    FGDEF_Node_atPoint = 38
    FGDEF_Node_atCurve = 39
    FGDEF_Node_atSurface = 40
    FGDEF_Node_atSolid = 41
    FGDEF_Elem_atPoint = 42
    FGDEF_Elem_atCurve = 43
    FGDEF_Elem_atSurface = 44
    FGDEF_Elem_atSolid = 45
    FGDEF_Load_byPoint = 46
    FGDEF_Load_byCurve = 47
    FGDEF_Load_bySurface = 48
    FGDEF_BCo_byPoint = 49
    FGDEF_BCo_byCurve = 50
    FGDEF_BCo_bySurface = 51
    FGDEF_Text_byColor = 52
    FGDEF_Point_byColor = 53
    FGDEF_Curve_byColor = 54
    FGDEF_Surface_byColor = 55
    FGDEF_Volume_byColor = 56
    FGDEF_Solid_byColor = 57
    FGDEF_CSys_byColor = 58
    FGDEF_Node_byColor = 59
    FGDEF_Elem_byColor = 60
    FGDEF_Prop_byColor = 61
    FGDEF_Matl_byColor = 62
    FGDEF_Text_byLayer = 63
    FGDEF_Point_byLayer = 64
    FGDEF_Curve_byLayer = 65
    FGDEF_Surface_byLayer = 66
    FGDEF_Volume_byLayer = 67
    FGDEF_Solid_byLayer = 68
    FGDEF_CSys_byLayer = 69
    FGDEF_Node_byLayer = 70
    FGDEF_Elem_byLayer = 71
    FGDEF_Prop_byLayer = 72
    FGDEF_Matl_byLayer = 73
    FGDEF_Solid_ID = 74
    FGDEF_Solid_byCurve = 75
    FGDEF_Solid_bySurface = 76
    FGDEF_Curve_onSolid = 77
    FGDEF_Surface_onSolid = 78
    FGDEF_Point_byProp = 79
    FGDEF_Curve_byProp = 80
    FGDEF_Surface_byProp = 81
    FGDEF_Volume_byProp = 82
    FGDEF_Solid_byProp = 83
    FGDEF_Contact_ID = 84
    FGDEF_Contact_byColor = 85
    FGDEF_Contact_byLayer = 86
    FGDEF_CSys_onNode = 87
    FGDEF_CSys_onPoint = 88
    FGDEF_Elem_byShape = 89
End Enum

Public Enum ElementType
    Rod = 1
    Bar = 2
    Tube = 3
    Link = 4
    Beam = 5
    Spring = 6
    DOFSpring =  7
    CurvedBeam = 8
    Gap = 9
    PlotOnly = 10
    ShearPanelLin = 11
    ShearPanelPara = 12
    MembraneLin = 13
    MembranePara = 14
    BendingOnlyLin = 15
    BendingOnlyPara = 16
    PlateLin = 17
    PlatePara = 18
    PlaneStrainLin = 19
    PlaneStrainPara = 20
    LaminateLin = 21
    LaminatePara = 22
    AxisymmetricLin = 23
    AxisymmetricPara = 24
    SolidLin = 25
    SolidPara = 26
    Mass = 27
    MassMatrix = 28
    Rigid = 29
    StiffnessMatrix = 30
    CurvedTube = 31
    PlotOnlyPlate = 32
    SlideLine = 33
    Contact = 34
End Enum

Public Enum ElementTopology
    Line = 0
    Tri3 = 2
    Tri6 = 3
    Quad4 = 4
    Quad8 = 5
    Tetra4 = 6
    Wedge6 = 7
    Brick8 = 8
    Point = 9
    Tetra10 = 10
    Wedge15 = 11
    Brick20 = 12
    Rigid = 13
    MultiList = 15
    Contact = 16
End Enum

Public Enum NodeType
    Node = 0
    Scalar = 1
    Extra = 2
End Enum

Public Enum MaterialType
    Iso = 0
    Ortho2D = 1
    Ortho3D = 2
    Aniso2D = 3
    Aniso3D = 4
    Hyperelastic = 5
    General = 6
    Fluid = 7
End Enum

Public Enum MaterialNonlinearType
    Linear = 0
    NonlinearElastic = 1
    Plastic = 2
    ElastoPlastic = 3
End Enum

Public Enum MaterialYieldCriterion
    vonMises = 0
    Tresca = 1
    Mohr_Coloumb = 2
    Drucker_Prager = 3
End Enum

Public Enum MaterialCreepType
    None = 0
    Empirical = 1
    Tabular = 2
End Enum

Public Enum PropertyType
    Rod = 1
    Bar = 2
    Tube = 3
    Link = 4
    Beam = 5
    Spring = 6
    DOFSpring = 7
    CurvedBeam = 8
    Gap = 9
    PlotOnly = 10
    ShearPanelLin = 11
    ShearPanelPara = 12
    MembraneLin = 13
    MembranePara = 14
    BendingOnlyLin = 15
    BendingOnlyPara = 16
    PlateLin = 17
    PlatePara = 18
    PlaneStrainLin = 19
    PlaneStrainPara = 20
    LaminateLin = 21
    LaminatePara = 22
    AxisymmetricLin = 23
    AxisymmetricPara = 24
    SolidLin = 25
    SolidPara = 26
    Mass = 27
    MassMatrix = 28
    Rigid = 29
    StiffnessMatrix = 30
    CurvedTube = 31
    PlotOnlyPlate = 32
    SlideLine = 33
    Contact = 34
End Enum

Public Enum BarBeamType
     RectangularBar = 1
     RectangularTube = 2
     TrapezoidalBar = 3
     TrapezoidalTube = 4
     CircularBar = 5
     CircularTube = 6
     HexBar = 7
     HexTube = 8
     I = 9
     Channel = 10
     Angle = 11
     T = 12
     Z = 13
     Hat = 14
     General = 15
End Enum

Public Enum CSysType
    Rectangular = 0
    Cylindrical = 1
    Spherical = 2
End Enum

Public Enum BCGeom_geomType
    Point = 3
    Curve = 4
    Surface = 5
End Enum

Public Enum BCGeomType
    General_6_DOF_Control_in_output_sys = 1
    Constrained_normal_to_surface = 2
    Constrained_in_all_in_surface_directions = 3
    Allow_sliding_in_a_direction = 4
    Cylinder_DOF_control_directions = 5
End Enum

Public Enum LoadSet_NLOn
    Off = 0
    Static = 1
    Creep = 2
    Transient = 3
End Enum

Public Enum LoadSet_NLConvergenceFlag
    Displacement = 0
    Load = 1
    Work = 2
End Enum

Public Enum LoadSet_NLSolutionOverride
    none_advanced = 0
    Full_Newton_Raphson = 1
    Modified_Newton_Raphson = 2
End Enum

Public Enum LoadSet_NLNewtRaphLineSearch
    Skip = 1
End Enum

Public Enum LoadSet_NLNewtRaphQuasiNewton
    Skip = 1
End Enum

Public Enum LoadSet_NLNewtRaphBisection
    Skip = 1
End Enum

Public Enum LoadSet_DYNOn
    Off = 0
    Direct = 1
    Modal = 2
End Enum

Public Enum LoadSet_DYNType
    Off = 0
    Transient = 1
    Freq = 2
End Enum

Public Enum LoadSet_DYNMassFormulation
    Default = 0
    Lumped = 1
    Coupled = 2
End Enum

Public Enum LoadSet_DYNDataRecovery
    ModeDisplacement = 0
    ModeAcceleration = 1
    Matrix = 2
End Enum

Public Enum LoadMesh_function
    LoadFunc_EmissivityFunc = 0
    Absorbtivity_vs_Temp = 1
    Temp_vs_Temp = 2
    ViewFactor_vs_Time = 3
    Phase_vs_Freq = 4
End Enum

Public Enum LoadMesh_LoadType
    nForce = 1
    nMoment = 2
    nDisplacement = 3
    nRotDisplacement = 4
    nVelocity = 5
    nRotVelocity = 6
    nAcceleration = 7
    nRotAcceleration = 8
    nHeatFlux = 10
    nHeatGen = 11
    Transient = 12
    nPressure = 13
    nTotalPressure = 14
    nScalar = 15
    nSteamQuality = 16
    nHumidity = 17
    nFluidHeight = 18
    nUnknownCondition = 19
    nSlipCondition = 20
    nFanCurve = 21
    nPeriodic = 22
    eLineLoad = 41
    ePressure = 42
    eHeatFlux = 44
    eConvection = 45
    eRadiation = 46
    eHeatGen = 47
End Enum

Public Enum LoadGeom_function
    LoadFunc_EmissivityFunc = 0
    Absorbtivity_vs_Temp = 1
    Temp_vs_Temp = 2
    ViewFactor_vs_Time = 3
    Phase_vs_Freq = 4
End Enum

Public Enum LoadGeom_dirmode
    None = 0
    Vector = 1
    AlongCurve = 2
    Normal2Plane = 3
    Normal2Surface = 4
End Enum

Public Enum LoadGeom_variation
    None_Constant = 0
    Equation = 1
    Function = 2
    Interpolation = 3
End Enum

Public Enum LoadGeom_LoadType
    pnForce = 81
    pnMoment = 82
    pnDisp = 83
    pnRotDisp = 84
    pnVelocity = 85
    pnRotVelocity = 86
    pnAccel = 87
    pnRotAccel = 88
    pnTemp = 89
    pnHeatFlux = 90
    pnHeatGen = 91
    pnPressure = 92
    pnTotalPressure = 93
    pnScalar = 94
    pnSteamQuality = 95
    pnHumidity = 96
    pnFluidHeight = 97
    pnUnknownCondition = 98
    pnSlipCondition = 99
    pnFanCurve = 100
    pnPeriodic = 101
    cnForce = 121
    cnForcePerLength = 122
    cnForceAtNode = 123
    cnMoment = 124
    cnMomentPerLength = 125
    cnMomentAtNode = 126
    cnDisp = 127
    cnRotDisp = 128
    cnVelocity = 129
    cnRotVelocity = 130
    cnAccel = 131
    cnRotAccel = 132
    cnTemp = 133
    cnHeatFlux = 134
    cnHeatFluxPerLength = 135
    cnHeatFluxAtNode = 136
    cnHeatGen = 137
    cePressure = 138
    ceTemp = 139
    ceHeatFlux = 140
    ceConvection = 141
    ceRadiation = 142
    ceHeatGen = 143
    cnPressure = 144
    cnTotalPressure = 145
    cnScalar = 146
    cnSteamQuality = 147
    cnHumidity = 148
    cnFluidHeight = 149
    cnUnknownCondition = 150
    cnSlipCondition = 151
    cnFanCurve = 152
    cnPeriodic = 153
    snForce = 161
    snForcePerArea = 162
    snForceAtNode = 163
    snMoment = 164
    snMomentPerArea = 165
    snMomentAtNode = 166
    snDisp = 167
    snRotDisp = 168
    snVelocity = 169
    snRotVelocity = 170
    snAccel = 171
    snRotAccel = 172
    snTemp = 173
    snHeatFlux = 174
    snHeatFluxPerArea = 175
    snHeatFluxAtNode = 176
    snHeatGen = 177
    sePressure = 178
    seTemp = 179
    seHeatFlux = 180
    seConvection = 181
    seRadiation = 182
    seHeatGen = 183
    snPressure = 184
    snTotalPressure = 185
    snScalar = 186
    snSteamQuality = 187
    snHumidity = 188
    snFluidHeight = 189
    snUnknownCondition = 190
    snSlipCondition = 191
    snFanCurve = 192
    snPeriodic = 193
End Enum

Public Enum Group_LayerMode
    All_Layers = 0
    Equal_or_Above_Max = 1
    Equal_or_Below_Min = 2
    Between = 3
    Outside = 4
    Single_Layer = 5
End Enum

Public Enum Group_CoordClipMode
    Greater = 0
    Less = 1
    Between = 2
    Outside = 3
End Enum

Public Enum Group_CoordClipDir
    X_or_R = 0
    Y_or_theta = 1
    Z_or_phi = 2
End Enum

Public Enum Group_PlaneClipMode
    Off = 0
    Screen = 1
    Plane = 2
    Volume = 3
End Enum

Public Enum Group_RangeType
    CSys_ID = 0
    CSys_byDefCSys = 1
    CSys_byType = 2
    Point_ID = 3
    Point_byDefCSys = 4
    Point_onCurve = 5
    Curve_ID = 6
    Curve_byPoint = 7
    Curve_onSurface = 8
    Surface_ID = 9
    Surface_byCurve = 10
    Surface_onVolume = 11
    Volume_ID = 12
    Volume_bySurface = 13
    Text_ID = 14
    Boundary_ID = 15
    Boundary_byCurve = 16
    Node_ID = 17
    Node_byDefCSys = 18
    Node_byOutCSys = 19
    Node_onElem = 20
    Elem_ID = 21
    Elem_byMatl = 22
    Elem_byProp = 23
    Elem_byType = 24
    Elem_byNode = 25
    Matl_ID = 26
    Matl_onProp = 27
    Matl_onElem = 28
    Matl_byType = 29
    Prop_ID = 30
    Prop_onElem = 31
    Prop_byMatl = 32
    Prop_byType = 33
    Load_byNode = 34
    Load_byElem = 35
    BCo_ID = 36
    BEq_byNode = 37
    Node_atPoint = 38
    Node_atCurve = 39
    Node_atSurface = 40
    Node_atSolid = 41
    Elem_atPoint = 42
    Elem_atCurve = 43
    Elem_atSurface = 44
    Elem_atSolid = 45
    Load_byPoint = 46
    Load_byCurve = 47
    Load_bySurface = 48
    BCo_byPoint = 49
    BCo_byCurve = 50
    BCo_bySurface = 51
    Text_byColor = 52
    Point_byColor = 53
    Curve_byColor = 54
    Surface_byColor = 55
    Volume_byColor = 56
    Solid_byColor = 57
    CSys_byColor = 58
    Node_byColor = 59
    Elem_byColor = 60
    Prop_byColor = 61
    Matl_byColor = 62
    Text_byLayer = 63
    Point_byLayer = 64
    Curve_byLayer = 65
    Surface_byLayer = 66
    Volume_byLayer = 67
    Solid_byLayer = 68
    CSys_byLayer = 69
    Node_byLayer = 70
    Elem_byLayer = 71
    Prop_byLayer = 72
    Matl_byLayer = 73
    Solid_ID = 74
    Solid_byCurve = 75
    Solid_bySurface = 76
    Curve_onSolid = 77
    Surface_onSolid = 78
    Point_byProp = 79
    Curve_byProp = 80
    Surface_byProp = 81
    Volume_byProp = 82
    Solid_byProp = 83
    Contact_ID = 84
    Contact_byColor = 85
    Contact_byLayer = 86
    CSys_onNode = 87
    CSys_onPoint = 88
    Elem_byShape = 89
End Enum

Public Enum Group_ListType
    CSys = 0
    Point = 1
    Curve = 2
    Surface = 3
    Volume = 4
    Text = 5
    Boundary = 6
    Node = 7
    Elem = 8
    Material = 9
    Property = 10
    Nodal_Load = 11
    Elem_Load = 12
    Constraint = 13
    Cosntraint_Equations = 14
    Point_Loads = 15
    Curve_Loads = 16
    Surface_Loads = 17
    Point_Constraints = 18
    Curve_Constraints = 19
    Surface_Const = 20
    Solids = 21
    Contact_Segments = 22
End Enum

Public Enum Group_Range_include
    Remove = 0
    Add = 1
    Exclude = -1
End Enum

Public Enum AnalysisSet_Solver
    Unknown = 0
    MSC_Nastran = 4
    ANSYS = 5
    ABAQUS = 16
    MSC_Marc = 29
End Enum

Public Enum AnalysisSet_AnalysisType
    Unknown = 0
    Static = 1
    Modes = 2
    Transient = 3
    Frequency_Response = 4
    Response_Spectrum = 5
    Random = 6
    Linear_Buckling = 7
    Design_Opt = 8
    Explicit = 9
    Nonlinear_Static = 10
    Nonlinear_Buckling = 11
    Nonlinear_Transient = 12
    Comp_Fluid_Dynamics = 19
    Steady_State_Heat_Transfer = 20
    Transient_Heat = 21
End Enum

Public Enum AnalysisSet_Output
    off = 0
    full_model = -1
End Enum

Public Enum AnalysisSet_Destination
    Default = 0
    Print = 1
    Post = 2
    Print_Post = 3
    Punch = 4
    Punch_Post = 5
End Enum

Public Enum AnalysisSet_Imaginary
    Magnitude_Phase = 0
    Real_Imaginary = 1
End Enum

Public Enum AnalysisSet_AbaHistStepAmp
    Default = 0
    Step = 1
    Ramp = 2
End Enum

Public Enum AnalysisSet_AbaHistStepLoad
    New = 0
    Modify = 1
End Enum

Public Enum AnalysisSet_AbaHistStepConstr
    New = 0
    Modify = 1
End Enum

Public Enum AnalysisSet_MarHistCtrlMethod
    None = 0
    Full_Newton_Raphson = 1
    Modified_Newton_Raphson = 2
    Strain_Correction_Newton_Raphson = 3
    Secant_Method = 8
End Enum

Public Enum AnalysisSet_MarHistSolverMeth
    None = 0
    Profile_Direct = 1
    Sparse_Iterative = 2
    Sparse_Direct = 3
    Hardware_Provided_Sparse_Direct = 4
    MultiFrontal_Direct_Sparse = 5
End Enum

Public Enum AnalysisSet_MarHistConvergeMeth
    Rel_Force_Residuals = 0
    Rel_Disp_Residuals = 1
    Rel_Strain_Ener_Res = 2
    Abs_Force_Residuals = 3
    Abs_Disp_Residuals = 4
    Abs_Strain_Ener_Res = 5
End Enum

Public Enum AnalysisSet_MarHistChecking
    None = 0
    Strum_Sequence = 1
End Enum

Public Enum AnalysisSet_MarHistAnalCaseSol
    Static = 1
    Normal_Modes = 2
    Buckling = 3
End Enum

Public Enum AnalysisSet_MarIncArcLenMeth
    None = 0
    Crisfield = 1
    Riks = 2
    Modified_Riks = 3
    Crisfield_Modified_Riks = 4
End Enum

Public Enum AnalysisSet_NasBulkLargeField
    Small = 0
    Large_Csys_Matl_Prop = 1
    Large_All_But_Elem = 2
    Large = 3
End Enum

Public Enum AnalysisSet_NasNonlinEpsFlag
    Temperature = 0
    Load = 1
    Work = 2
End Enum

Public Enum AnalysisSet_NasModeMethod
    Givens = 0
    Modified_Givens = 1
    Inverse_Power = 2
    Inverse_Power_Sturm = 3
    Householder = 4
    Modified_Householder = 5
    Lanczos = 6
End Enum

Public Enum AnalysisSet_NasModeSolutionType
    Direct = 1
    Modal = 2
End Enum

Public Enum AnalysisSet_NasModeNormOpt
    Mass = 0
    Max = 1
    Point = 2
End Enum

Public Enum AnalysisSet_NasModeMassForm
    Default = 0
    Lumped = 1
    Coupled = 2
End Enum

Public Enum AnalysisSet_NasAppSpecMethod
    ABS = 0
    SRSS = 1
    NRL = 2
    NRLO = 3
End Enum

Public Enum AnalysisSet_NasGCheckMsg
    Fatal = 0
    Inform = 1
    Warn = 2
End Enum

Public Enum AnalysisSet_NasMCheckWtUnits
    MASS = 0
    WEIGHT = 1
End Enum

Public Enum AnalysisSet_MarModFolOpt
    Off_Incremental = 1
    Last_Iter_Inc = 2
    Beg_Disp_Incr = 3
    Off_Total = 4
    Last_Iter_Total = 5
    Beg_Disp_To = 6
End Enum

Public Enum AnalysisSet_MarModPlasOpt
    Add_Normal_Small_Strain = 1
    Add_Radial_Small_Strain = 2
    Add_Normal_Large_Strain = 3
    Add_Radial_Large_Strain = 4
    Multiplicative = 5
End Enum

Public Enum View_Mode
    Draw = 0
    Feature = 1
    Quick_Hide = 2
    Hide = 3
    Free_Edge = 4
    Free_Face = 5
    XY_vs_ID = 6
    XY_vs_SET = 7
    XY_vs_VALUE = 8
    XY_vs_POSITION = 9
    XY_of_Function = 10
End Enum

Public Enum View_Deformed
    Off = 0
    Deformed = 1
    Animate = 2
    Animate_MultiCase = 3
    Arrow = 4
End Enum

Public Enum View_Contour
    Off = 0
    Contour = 1
    Criteria = 2
    Beam_Diagram = 3
    IsoSurface = 4
    Section_Cut = 5
End Enum

Public Enum View_OptionTypes_Index
    ' Labels, Entities and Colors
    Label_Parameters = 0
    Coordinate_System = 1
    Point = 2
    Curve = 3
    Curve_Mesh_Size = 24
    Surface = 4
    Volume = 5
    Text = 6
    Boundary = 27
    Node = 7
    Node_Perm_Constraint = 8
    Element = 9
    Element_Directions = 10
    Element_Offsets_Releases = 11
    Element_Orientation_Shape = 12
    Element_Beam_Y_Axis = 13
    Load_Vectors = 78
    Load_Force = 14
    Load_Moment = 15
    Load_Thermal = 16
    Load_Distributed = 71
    Load_Pressure = 17
    Load_Acceleration = 18
    Load_Velocity = 72
    Load_Enforced_Displacement = 19
    Load_Nonlinear_Force = 73
    Load_Heat_Generation = 66
    Load_Heat_Flux = 67
    Load_Convection = 68
    Load_Radiation = 69
    Load_Fluid_Tracking = 81
    Load_Unknown_Condition = 82
    Load_Slip_Wall_Condition = 83
    Load_Fan_Curve = 84
    Load_Periodic_Condition = 85
    Constraint = 20
    Constraint_Equation = 21
    Contact_Segment = 79
    ' Tools and View Style
    Free_Edge_and_Face = 22
    Shrink_Elements = 23
    Fill_Backfaces_and_Hidden = 25
    Filled_Edges = 26
    Render_Options = 77
    Shading = 28
    Perspective = 29
    Stereo = 30
    View_Legend = 31
    View_Axes = 32
    Origin = 33
    Workplane_and_Rulers = 34
    Workplane_Grid = 35
    Clipping_Planes = 36
    Symbols = 37
    View_Aspect_Ratio = 38
    Curve_and_Surface_Accuracy = 39
    ' PostProcessing
    Post_Titles = 40
    Deformed_Style = 41
    Vector_Style = 42
    Animated_Style = 43
    Deformed_Model = 44
    Undeformed_Model = 45
    Trace_Style = 74
    Contour_Criteria_Style = 46
    Contour_Criteria_Levels = 47
    Contour_Criteria_Legend = 48
    Criteria_Limits_Beam_Diagrams = 49
    Criteria_Elements_that_Pass = 50
    Criteria_Elements_that_Fail = 51
    Isosurface = 76
    Contour_Vector_Style = 75
    XY_Titles = 52
    XY_Legend = 53
    XY_Axes_Style = 54
    XY_X_Range_Grid = 55
    XY_Y_Range_Grid = 56
    XY_Curve_1 = 57
    XY_Curve_2 = 58
    XY_Curve_3 = 59
    XY_Curve_4 = 60
    XY_Curve_5 = 61
    XY_Curve_6 = 62
    XY_Curve_7 = 63
    XY_Curve_8 = 64
    XY_Curve_9 = 65
End Enum

Public Enum Report_DataType
    Nodal_Results = 7
    Elemental_Results = 8
End Enum

Public Enum Function_type
    Dimensionless = 0
    vs_Time = 1
    vs_Temp = 2
    vs_Freq = 3
    vs_Stress = 4
    Func_vs_Temp = 5
    Struct_Damp_vs_Freq = 6
    Crit_Damp_vs_Freq = 7
    Q_Damp_vs_Freq = 8
    vs_Strain_Rate = 9
    Func_vs_Strain_Rate = 10
    vs_Curve_Length = 11
    vs_Curve_Param = 12
    Stress_vs_Strain = 13
    Stress_vs_Plastic_Strain = 14
    Function_vs_Value = 15
    Function_vs_Critical_Damping = 16
End Enum

Public Enum Optim_type
    Goal = 1
    Vary = 2
    Limit = 3
End Enum

Public Enum Optim_goal
    None = 0
    MinWeight = 1
End Enum

Public Enum Optim_vary
    None = 0
    RodArea = 1
    RodTorsion = 2
    BarArea = 3
    BarI1 = 4
    BarI2 = 5
    BarTorsion = 6
    PlateThickness = 7
End Enum

Public Enum Optim_limit
    None = 0
    NodXDisp = 1
    NodYDisp = 2
    NodZDisp = 3
    NodXRDisp = 4
    NodYRDisp = 5
    NodZRDisp = 6
    RodAxialStress = 7
    RodTorsionStress = 8
    RodAxialStrain = 9
    RodTorsionStrain = 10
    BarAxialStress = 11
    BarMaxStress = 12
    BarMinStress = 13
    BarAxialStrain = 14
    BarMaxStrain = 15
    BarMinStrain = 16
    PltXNormalStress = 17
    PltYNormalStress = 18
    PltXYShearStress = 19
    PltMaxPrinStress = 20
    PltMinPrinStress = 21
    PltVonMisesStress = 22
    PltXNormalStrain = 23
    PltYNormalStrain = 24
    PltXYShearStrain = 25
    PltMaxPrinStrain = 26
    PltMinPrinStrain = 27
    PltVonMisesStrain = 28
End Enum

Public Enum Optim_varyType
    Property = 11
End Enum

Public Enum Optim_respType
    Node = 7
    Property = 11
End Enum

Public Enum OutputSet_program
    Unknown = 0
    MSC_N4W_Generated = 1
    PAL = 2
    PAL_2 = 3
    MSC_NASTRAN = 4
    ANSYS = 5
    STARDYNE = 6
    COSMOS = 7
    PATRAN = 8
    FEMAP_Neutral = 9
    ALGOR = 10
    SSS_NASTRAN = 11
    Comma_Separated = 12
    UAI_NASTRAN = 13
    Cosmic_NASTRAN = 14
    STAAD = 15
    ABAQUS = 16
    WECAN = 17
    MTAB_SAP = 18
    CDA_Sprint = 19
    CAEFEM = 20
    I_DEAS = 21
    ME_NASTRAN = 22
    CSA_NASTRAN = 26
    CFDesign = 28
    LS_DYNA = 31
    MARC = 32
    SINDA = 33
End Enum

Public Enum OutpuSet_analysis
Unknown = 0
Static = 1
Modes = 2
Transient = 3
Frequency_Response = 4
Response_Spectrum = 5
Random = 6
Linear_Buckling = 7
Design_Opt = 8
Explicit = 9
Nonlinear_Static = 10
Nonlinear_Buckling = 11
Nonlinear_Transient = 12
Comp_Fluid_Dynamics = 19
Steady_State_Heat_Transfer = 20
Transient_Heat = 21
End Enum

Public Enum Output_category
Any = 0
Disp = 1
Accel = 2
Force = 3
Stress = 4
Strain = 5
Temp = 6
End Enum

Public Enum Output_location
Nodal = 7
Elemental_Centroid_or_Element_Corner = 8
End Enum

Public Enum Point_type
    Default = 0
        Solid = 1
    End Enum

    Public Enum Point_engine
    None = 0
    Parasolid = 1
    ACIS = 2
End Enum

Public Enum Curve_type
Line = 0
Arc = 1
Circle = 2
Spline = 3
BSpline = 4
Solid = 5
End Enum

Public Enum Curve_attrOrientType
Orient_By_Vector = 0
Orient_By_Location = 1
Orient_By_Vector_Reversed_ElDir = 2
Orient_By_Location_Reversed_ElDir = 3
End Enum

Public Enum Curve_attrOffsetType
Offset_by_Vector = 0
Offset_Radial = 1
Offset_by_Location = 2
End Enum

Public Enum Curve_Engine
None = 0
Parasolid = 1
ACIS = 2
End Enum

Public Enum Curve_biasMethod
Equal_length_spacing = 0
Linear_Bias = 1
Geometric_Bias = 2
End Enum

Public Enum Curve_biasLoc
Small_elements_at_end_of_curve = 1
Small_elements_at_center_of_curve = 2
Small_elements_at_both_ends_of_curve = 3
End Enum

Public Enum Surface_type
Bilinear = 0
Ruled = 1
Revolution = 2
Coons = 3
Bezier = 4
Solid = 5
Not_Used = 6
Boundary = 7
End Enum

Public Enum Surface_approach
Not_Specified = 0
Free_Parametric = 1
Free_Planar = 2
Mapped_Four_Corner = 3
Mapped_Three_Corner = 4
Mapped_Three_Corner_Fan = 5
Link_to_Surface = 6
Fast_Tri_Parametric = 7
Fast_Tri_Planar = 8
End Enum

Public Enum Surface_Engine
None = 0
Parasolid = 1
ACIS = 2
End Enum

Public Enum Surface_BoundaryMode
Planar_Interpolate_Curves = 0
Map_to_Surface = 3
End Enum

Public Enum Solid_Type
Volume = 6
Solid = 39
End Enum

Public Enum Solid_VolType
Brick = 0
Wedge = 1
Pyramid = 2
Tetra = 3
End Enum

Public Enum Solid_Engine
None = 0
Parasolid = 1
ACIS = 2
End Enum

Public Enum Event_wParam
FEVENT_INITIALIZE = 1
FEVENT_NEWMODEL = 2
FEVENT_ENDMODEL = 3
FEVENT_SHUTDOWN = 4
FEVENT_COMMAND = 5
End Enum

Public Enum Event_lParam
View_Redraw  = 2001
View_Regenerate  = 2002
Tools_Undo  = 1101
Tools_Redo  = 1102
File_Save  = 1003
File_Open  = 1002
File_Exit  = 1025
End Enum

Public Enum PgSetup_VertAlign
Center = 0
Top = 1
Bottom = 2
End Enum

Public Enum PgSetup_HorzAlign
Center = 0
Left = 1
Right = 2
End Enum

Public Enum Solid_UpdateResizeMode
Keep_existing_sizes = 0
Update_sizes_on_curves_that_exceed_a_certain_length_change_tolerance = 1
Resize_all = 2
End Enum

Public Enum Pref_DBLowDiskWarning
Disable = 0
End Enum

Public Enum Pref_GeomEngine
Std = 0
Parasolid = 1
ACIS = 2
End Enum

Public Enum Pref_InterfaceStyle
Structural = 0
Thermal = 1
Adv_Thermal = 2
End Enum

Public Enum Pref_PosCommandToolbar
Top = 1
Left = 2
Right = 3
End Enum

Public Enum Pref_PosViewToolbar
Top = 1
Left = 2
Right = 3
End Enum

Public Enum Pref_ViewDynamicMode
Fast_Redraw = 0
Reduced_Bitmap = 1
Full_Bitmap = 2
End Enum

Public Enum Info_NodeType
Node = 0
Scalar = 1
Extra = 2
End Enum

Public Enum Info_SnapTo
Screen = 0
Grid = 1
Point = 2
Node = 3
End Enum

Public Enum Info_SnapStyle
Invisible = 0
Dots = 1
Lines = 2
End Enum

Public Enum Info_SnapSpacingMode
Automatic = 0
Uniform = 1
Nonuniform = 2
End Enum

Public Enum Info_MatlAngleMethod
    None_Off = 0
    Vector = 1
    CoordSys = 2
    Angle = 3
End Enum

Public Enum Info_MatlAngleDir
    X = 0
    Y = 1
    Z = 2
End Enum

Public Enum Info_GroupAutomaticAdd
    Off = 0
    Active_Group = -1
End Enum

' vim: set ft=vb:
