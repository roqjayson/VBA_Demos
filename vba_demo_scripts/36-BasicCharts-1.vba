Option Explicit

Sub CreateChart()

    Dim Chrt As ChartObject
    Dim DataRng As Range
    
    
    
    ' Add chart object
    
    Set Chrt = ActiveSheet.ChartObjects.Add(Left:=400, _
                                            Width:=400, _
                                            Height:=400, _
                                            Top:=50)
                                                               
   
                                        
    'Add Data
    
    Set DataRng = Range("B3:E7")
    Chrt.Chart.SetSourceData Source:=DataRng ' Default clustered bar
    
    ' Add chart type
    Chrt.Chart.ChartType = xlBarClustered
    
    ' Add Title
    Chrt.Chart.HasTitle = True ' Default Chart Title
    
    ' Create a reference for the title
    
    Dim chrtTitle As ChartTitle
    Set chrtTitle = Chrt.Chart.ChartTitle
    
    ' Formatting title
        With chrtTitle
            .Text = "Performance"
            .Shadow = False
            .Characters.Font.Bold = False
            .Characters.Font.Name = "Arial Nova"
        End With
    
    ' Add legends
    
    Chrt.Chart.HasLegend = True
    
    ' Reference to the legend
    
    Dim chrtLeg As Legend
    Set chrtLeg = Chrt.Chart.Legend
    
        chrtLeg.Position = xlLegendPositionTop
        chrtLeg.Height = 40
        
    ' Remove gridlines
    Chrt.Chart.SetElement msoElementPrimaryCategoryGridLinesNone
    Chrt.Chart.SetElement msoElementPrimaryValueGridLinesNone
    
    ' Axes
    
    Chrt.Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chrt.Chart.Axes(xlValue, xlPrimary).HasTitle = True
    
    ' Reference to single axis
    
    Dim axCatTitle As axistitle
    Set axCatTitle = Chrt.Chart.Axes(xlCategory, xlPrimary).axistitle
        
        ' Format axes
        
        axCatTitle.Text = "Years"
        axCatTitle.HorizontalAlignment = xlCenter
        axCatTitle.Characters.Font.Color = vbRed
        
    ' Reference to another axis
    Dim axValTitle As axistitle
    Set axValTitle = Chrt.Chart.Axes(xlValue, xlPrimary).axistitle
        
        ' Format axes
        
        axValTitle.Text = "Currency USD"
        axValTitle.HorizontalAlignment = xlCenter
        axValTitle.Characters.Font.Color = vbRed
    
    Dim axValLabel As Axis
    Set axValLabel = Chrt.Chart.Axes(xlValue)
    
        axValLabel.TickLabels.NumberFormat = "$#0,K"
        
    
    
    

End Sub


