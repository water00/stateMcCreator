# stateMcCreator
Creates state machine diagram and state machine code from worksheet

# How To
The excel sheet stateMcCreator.xlsm file contains VB script. Follwoing are the steps and notes

1. Enter your states on the blue line. You can add more columns. All states should be single word (C identifier)
2. Enter your events on the green line. You cna add more rows. All Events have to be single word too.
3. Enter your actions on the matrix between states and events. Each action should be in a seprate line (Use Alt Enter to create a new line within cell)
4. If state changes to next state, then nextState = <StateName> should be the last line in the cell. StateName should match state names defines in blue line
5. Pressing Generate Image would generate the state image at the bottom of the matrix.
6. Please note previous image would be replaced by new image
7. Pressing Generate Code will generate C++ code with state changes.
8. Please note previous code will be replaced by new code



# Tutorials and EXamples
Couple of exampl Excel sheets are in TEX directory. 
