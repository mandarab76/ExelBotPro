/* global Office, Excel */

// Office.js initialization
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        console.log('ExcelBot Pro loaded successfully');
        initializeEventHandlers();
    }
});

// Initialize all event handlers
function initializeEventHandlers() {
    // VBA Macro Generation
    document.getElementById('generateMacroBtn').addEventListener('click', generateVBAMacro);
    document.getElementById('copyMacroBtn').addEventListener('click', copyMacroToClipboard);
    document.getElementById('insertMacroBtn').addEventListener('click', insertMacroIntoVBA);
    
    // Workbook Analysis
    document.getElementById('analyzeWorkbookBtn').addEventListener('click', analyzeWorkbook);
    
    // GitHub Integration
    document.getElementById('pushToGithubBtn').addEventListener('click', pushToGitHub);
    
    // Quick Actions
    document.getElementById('formatRangeBtn').addEventListener('click', formatSelectedRange);
    document.getElementById('createChartBtn').addEventListener('click', createChart);
    document.getElementById('sortDataBtn').addEventListener('click', sortData);
    document.getElementById('filterDataBtn').addEventListener('click', filterData);
}

// Generate VBA Macro based on user description
async function generateVBAMacro() {
    const taskDescription = document.getElementById('taskDescription').value.trim();
    
    if (!taskDescription) {
        showError('Please enter a task description');
        return;
    }
    
    try {
        const macroCode = generateMacroCode(taskDescription);
        document.getElementById('macroCode').textContent = macroCode;
        document.getElementById('macroOutput').style.display = 'block';
    } catch (error) {
        showError('Error generating macro: ' + error.message);
    }
}

// Generate actual VBA code based on task description
function generateMacroCode(taskDescription) {
    // Enhanced macro generation with more specific patterns
    const lowerTask = taskDescription.toLowerCase();
    
    // Pattern matching for common tasks
    if (lowerTask.includes('highlight') || lowerTask.includes('color')) {
        return `Sub HighlightCells()
    ' Task: ${taskDescription}
    Dim rng As Range
    Dim cell As Range
    
    ' Get the selected range or used range
    If Selection.Count > 1 Then
        Set rng = Selection
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    
    ' Loop through each cell
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value > 100 Then
                cell.Interior.Color = RGB(255, 200, 200) ' Light red
            ElseIf cell.Value > 50 Then
                cell.Interior.Color = RGB(255, 255, 200) ' Light yellow
            End If
        End If
    Next cell
    
    MsgBox "Highlighting completed!", vbInformation
End Sub`;
    } else if (lowerTask.includes('sort') || lowerTask.includes('order')) {
        return `Sub SortData()
    ' Task: ${taskDescription}
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Define the range to sort
    Set sortRange = ws.Range("A1:Z" & lastRow)
    
    ' Sort by first column
    sortRange.Sort Key1:=ws.Range("A1"), _
                    Order1:=xlAscending, _
                    Header:=xlYes
    
    MsgBox "Data sorted successfully!", vbInformation
End Sub`;
    } else if (lowerTask.includes('filter')) {
        return `Sub FilterData()
    ' Task: ${taskDescription}
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Turn on AutoFilter
    If Not ws.AutoFilterMode Then
        ws.Range("A1").AutoFilter
    End If
    
    ' Apply filter (example: filter column A for values > 100)
    ws.Range("A1:A" & lastRow).AutoFilter Field:=1, Criteria1:=">100"
    
    MsgBox "Filter applied!", vbInformation
End Sub`;
    } else if (lowerTask.includes('chart') || lowerTask.includes('graph')) {
        return `Sub CreateChart()
    ' Task: ${taskDescription}
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set ws = ActiveSheet
    Set dataRange = Selection
    
    ' Create a new chart
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=50, Width:=400, Height:=300)
    
    With chartObj.Chart
        .SetSourceData Source:=dataRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Data Chart"
    End With
    
    MsgBox "Chart created successfully!", vbInformation
End Sub`;
    } else if (lowerTask.includes('format') || lowerTask.includes('style')) {
        return `Sub FormatRange()
    ' Task: ${taskDescription}
    Dim rng As Range
    
    Set rng = Selection
    
    With rng
        .Font.Name = "Arial"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Format header row
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(68, 114, 196)
        .Rows(1).Font.Color = RGB(255, 255, 255)
    End With
    
    MsgBox "Formatting applied!", vbInformation
End Sub`;
    } else {
        // Generic macro template
        return `Sub AutoMacro()
    ' Task: ${taskDescription}
    ' This is a template macro. Customize it based on your needs.
    
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = ActiveSheet
    Set rng = Selection
    
    ' Add your custom code here
    MsgBox "Processing: ${taskDescription}", vbInformation
    
    ' Example: Loop through selected cells
    Dim cell As Range
    For Each cell In rng
        ' Process each cell
        Debug.Print cell.Address & ": " & cell.Value
    Next cell
    
    MsgBox "Macro completed!", vbInformation
End Sub`;
    }
}

// Copy macro to clipboard
async function copyMacroToClipboard() {
    const macroCode = document.getElementById('macroCode').textContent;
    
    try {
        await navigator.clipboard.writeText(macroCode);
        showSuccess('Macro copied to clipboard!');
    } catch (error) {
        showError('Failed to copy: ' + error.message);
    }
}

// Insert macro into VBA Editor (informational only)
async function insertMacroIntoVBA() {
    showInfo('To insert this macro into Excel:\n\n1. Press Alt+F11 to open VBA Editor\n2. Insert > Module\n3. Paste the copied macro code\n4. Press Alt+F11 to return to Excel\n5. Run the macro from Developer > Macros');
}

// Analyze current workbook
async function analyzeWorkbook() {
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load('items/name');
            
            await context.sync();
            
            // Get active sheet info
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = activeSheet.getUsedRange();
            usedRange.load('rowCount, columnCount');
            
            await context.sync();
            
            // Build analysis content safely using DOM methods
            const analysisContent = document.getElementById('analysisContent');
            analysisContent.textContent = ''; // Clear previous content
            
            // Create structured content
            const createBold = (text) => {
                const bold = document.createElement('strong');
                bold.textContent = text;
                return bold;
            };
            
            const addLine = (parent, text, isBold = false) => {
                if (isBold) {
                    parent.appendChild(createBold(text));
                } else {
                    parent.appendChild(document.createTextNode(text));
                }
                parent.appendChild(document.createElement('br'));
            };
            
            addLine(analysisContent, 'Workbook Information:', true);
            analysisContent.appendChild(document.createElement('br'));
            
            analysisContent.appendChild(createBold('Total Sheets: '));
            analysisContent.appendChild(document.createTextNode(sheets.items.length.toString()));
            analysisContent.appendChild(document.createElement('br'));
            
            addLine(analysisContent, 'Sheet Names:', true);
            
            sheets.items.forEach((sheet, index) => {
                analysisContent.appendChild(document.createTextNode(`${index + 1}. ${sheet.name}`));
                analysisContent.appendChild(document.createElement('br'));
            });
            
            analysisContent.appendChild(document.createElement('br'));
            addLine(analysisContent, 'Active Sheet Data:', true);
            
            analysisContent.appendChild(document.createTextNode(`Rows: ${usedRange.rowCount}`));
            analysisContent.appendChild(document.createElement('br'));
            
            analysisContent.appendChild(document.createTextNode(`Columns: ${usedRange.columnCount}`));
            analysisContent.appendChild(document.createElement('br'));
            
            document.getElementById('analysisOutput').style.display = 'block';
        });
    } catch (error) {
        showError('Error analyzing workbook: ' + error.message);
    }
}

// Push macro to GitHub
async function pushToGitHub() {
    const token = document.getElementById('githubToken').value.trim();
    const repoName = document.getElementById('repoName').value.trim();
    const fileName = document.getElementById('fileName').value.trim();
    const macroCode = document.getElementById('macroCode').textContent;
    
    if (!token || !repoName || !fileName) {
        showError('Please fill in all GitHub fields');
        return;
    }
    
    if (!macroCode) {
        showError('Please generate a macro first');
        return;
    }
    
    try {
        // Note: This is a simplified example. In production, you'd use GitHub API
        showInfo('GitHub Integration:\n\nThis feature requires a backend service to securely handle GitHub authentication. For manual push:\n\n1. Go to github.com/' + repoName + '\n2. Create a new file: ' + fileName + '\n3. Paste the generated macro code\n4. Commit the changes');
        
        document.getElementById('githubOutput').style.display = 'block';
    } catch (error) {
        showError('GitHub push failed: ' + error.message);
    }
}

// Quick Action: Format Selected Range
async function formatSelectedRange() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.format.font.name = "Arial";
            range.format.font.size = 11;
            range.format.borders.getItem('EdgeTop').style = 'Continuous';
            range.format.borders.getItem('EdgeBottom').style = 'Continuous';
            range.format.borders.getItem('EdgeLeft').style = 'Continuous';
            range.format.borders.getItem('EdgeRight').style = 'Continuous';
            
            await context.sync();
            showSuccess('Range formatted successfully!');
        });
    } catch (error) {
        showError('Error formatting range: ' + error.message);
    }
}

// Quick Action: Create Chart
async function createChart() {
    try {
        await Excel.run(async (context) => {
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            const selectedRange = context.workbook.getSelectedRange();
            
            const chart = activeSheet.charts.add(Excel.ChartType.columnClustered, selectedRange, Excel.ChartSeriesBy.auto);
            chart.title.text = "Data Chart";
            chart.setPosition("E2", "L15");
            
            await context.sync();
            showSuccess('Chart created successfully!');
        });
    } catch (error) {
        showError('Error creating chart: ' + error.message);
    }
}

// Quick Action: Sort Data
async function sortData() {
    try {
        await Excel.run(async (context) => {
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = activeSheet.getUsedRange();
            
            usedRange.sort.apply([
                {
                    key: 0,
                    ascending: true
                }
            ], true, true, Excel.SortOrientation.rows);
            
            await context.sync();
            showSuccess('Data sorted successfully!');
        });
    } catch (error) {
        showError('Error sorting data: ' + error.message);
    }
}

// Quick Action: Filter Data
async function filterData() {
    try {
        await Excel.run(async (context) => {
            const activeSheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = activeSheet.getUsedRange();
            
            usedRange.autoFilter.apply(usedRange, 0, {
                filterOn: Excel.FilterOn.values
            });
            
            await context.sync();
            showSuccess('AutoFilter applied successfully!');
        });
    } catch (error) {
        showError('Error applying filter: ' + error.message);
    }
}

// Helper function to show success messages
function showSuccess(message) {
    const outputDiv = document.createElement('div');
    outputDiv.className = 'info-box success';
    outputDiv.textContent = message;
    
    // Add to analysis section temporarily
    const analysisContent = document.getElementById('analysisContent');
    analysisContent.textContent = '';  // Clear previous content
    analysisContent.appendChild(outputDiv);
    document.getElementById('analysisOutput').style.display = 'block';
    
    setTimeout(() => {
        document.getElementById('analysisOutput').style.display = 'none';
    }, 3000);
}

// Helper function to show error messages
function showError(message) {
    const outputDiv = document.createElement('div');
    outputDiv.className = 'info-box error';
    outputDiv.textContent = message;
    
    const analysisContent = document.getElementById('analysisContent');
    analysisContent.textContent = '';  // Clear previous content
    analysisContent.appendChild(outputDiv);
    document.getElementById('analysisOutput').style.display = 'block';
    
    setTimeout(() => {
        document.getElementById('analysisOutput').style.display = 'none';
    }, 3000);
}

// Helper function to show info messages
function showInfo(message) {
    const githubStatus = document.getElementById('githubStatus');
    githubStatus.className = 'info-box';
    
    // Clear previous content
    githubStatus.textContent = '';
    
    // Split by newlines and create text nodes with breaks
    const lines = message.split('\n');
    lines.forEach((line, index) => {
        githubStatus.appendChild(document.createTextNode(line));
        if (index < lines.length - 1) {
            githubStatus.appendChild(document.createElement('br'));
        }
    });
}
