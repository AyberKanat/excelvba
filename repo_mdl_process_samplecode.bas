Option Explicit

Sub UpdateOnTick()

'Declare the ticker variable
Dim lTicker As Long 'Counter variable for timing loop
Dim lTotalTicks As Long 'Number of timing ticks in simulation

'Declare collections and objects within collections
Dim Customers As Collection
Dim Stations As Collection
Dim objC As CCustomer
Dim objS As CStation

'Declare loop counter, random, and processing time variables
Dim lStCntr As Long
Dim lCustIDCntr As Long
Dim lStaTime As Long
Dim sngRand As Single
Dim lNextArrival As Long
Dim lCreateNextAt As Long

'New variables for multiple runs and run loop counter
Dim lNmbrRuns As Long
Dim lRunCntr As Long

'Declare variable for number of stations
Dim lNmbrStas As Long

'Assign value for number of stations
lNmbrStas = Worksheets("SimSetup").Range("C2").Value

'Create Customers, Stations, and WaitingRoom collections
Set Customers = New Collection
Set Stations = New Collection

'Create station objects and initialize values
Worksheets("SimSetup").Activate
Range("C4").Activate

For lStCntr = 1 To lNmbrStas

    'Add a station
    Set objS = New CStation
    Stations.Add objS
    
    'Assign Station ID and read Mean and StdDev from worksheet
    Stations(lStCntr).StaID = lStCntr
    Stations(lStCntr).StaMean = ActiveCell.Offset(lStCntr, 0)
    Stations(lStCntr).StaSD = ActiveCell.Offset(lStCntr, 1)
    Stations(lStCntr).NextSta = ActiveCell.Offset(lStCntr, 2)
    
    'Set StaIsIdle to 1
    Stations(lStCntr).StaIsIdle = 1
    
Next

'
'
'Set total ticks for each run
lTotalTicks = 2880 'Each tick represents 10 seconds in an eight-hour day

lCreateNextAt = 1

'Set CustID counter to 1
lCustIDCntr = 1


'Activate the SimSetup worksheet so you can read values from it
Worksheets("SimSetup").Activate

'Start the simulation!

'Outer For Next loop runs for the total ticks set above
For lTicker = 1 To lTotalTicks

'Check if a customer is due to be created on this tick
'If not, do nothing.
  If lTicker = lCreateNextAt Then
  
    Set objC = New CCustomer
    Customers.Add objC
    Customers(lCustIDCntr).CustID = lCustIDCntr
    Customers(lCustIDCntr).StartTime = 1
    Customers(lCustIDCntr).NextSta = 1
    Customers(lCustIDCntr).IsIdle = 1
    Customers(lCustIDCntr).IdleTime = 0
    Customers(lCustIDCntr).Entered = lTicker
    
    ActiveSheet.Calculate
    lCreateNextAt = lCreateNextAt + ActiveSheet.Range("K2").Value
    
    lCustIDCntr = lCustIDCntr + 1
  
  End If

'Start checking each customer's status. If they are idle,
'check which station they are in. If it's -1, they are done.
'If it's any other station, generate a processing time and
'set the customer's IsIdle property to 0.
    For Each objC In Customers
    
        If objC.Station <> -1 Then
        
          If objC.IsIdle = 1 Then
            
            'If the customer remains idle for this tick, add 1 to idle time.
            If Stations(objC.NextSta).StaIsIdle = 0 Then
                        
                objC.IdleTime = objC.IdleTime + 1
    
              'If not, set as not idle and generate processing time.
              Else:
                  objC.Station = objC.NextSta
                  
                  Stations(objC.Station).StaIsIdle = 0
                  
                  objC.NextSta = Stations(objC.Station).NextSta
                  sngRand = Rnd()
                  objC.EndTime = lTicker + Application.WorksheetFunction.Norm_Inv(sngRand, Stations(objC.Station).StaMean, _
                    Stations(objC.Station).StaSD)
                    
                  'Make sure the customer spends at least one tick in a station
                  If objC.EndTime <= lTicker Then objC.EndTime = lTicker + 1
                  
                  objC.IsIdle = 0
                  objC.StartTime = lTicker
                
            End If
            
          'If a customer's EndTime = the current tick, update Idle status to true
          'and move it to the next station
          ElseIf objC.EndTime = lTicker Then
            
                Stations(objC.Station).StaIsIdle = 1
                objC.IsIdle = 1
                objC.Station = objC.NextSta
                
                'If the customer's station is -1, it is done.
                'Set station to -1 and record the time it left the system.
                If objC.Station = -1 Then objC.Left = lTicker

        End If
    
      Else
      
        Stations(lNmbrStas).StaIsIdle = 1 'Sets last station to idle when a customer leaves it
        
    End If

    Next 'objC
    
Next lTicker


'Write results to Results worksheet
Worksheets("Results").Activate
ActiveSheet.Range("A1").Activate

'Step through each customer record and write property values
For Each objC In Customers

    ActiveCell.Offset(objC.CustID, 0).Value = objC.CustID
    ActiveCell.Offset(objC.CustID, 1).Value = objC.Entered
    ActiveCell.Offset(objC.CustID, 2).Value = objC.Left
    ActiveCell.Offset(objC.CustID, 3).Value = objC.Station
    ActiveCell.Offset(objC.CustID, 4).Value = objC.IsIdle
    ActiveCell.Offset(objC.CustID, 5).Value = objC.IdleTime

Next

End Sub
