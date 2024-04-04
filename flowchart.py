from graphviz import Digraph

dot = Digraph(comment='TM Creation Decision Process')

# Adjusting node shapes and streamlining the flow
dot.node('A', 'Start', shape='ellipse')
dot.node('B', 'Is it a new LCI project?', shape='diamond')
dot.node('C', 'Create Bilingual Glossary without TM\n[End]', shape='rectangle')
dot.node('D', 'Does the client provide a TM?', shape='diamond')
dot.node('E', 'Use client\'s TM to create bilingual glossary for Phrase\n[End]', shape='rectangle')
dot.node('F', 'Is it an existing LCI project without a TM?', shape='diamond')
dot.node('G', 'Create TM using last year\'s galley\nDecide on TM creation as assignment is confirmed\n[End]', shape='rectangle')
dot.node('H', 'Outsource Translation?', shape='diamond')
dot.node('I', 'Decide on TM creation based on manuscript availability\n[End]', shape='rectangle')

# Connecting nodes with revised logic to avoid unnecessary arrows
dot.edge('A', 'B')
dot.edge('B', 'C', label='Yes')
dot.edge('B', 'D', label='No')
dot.edge('D', 'E', label='Yes')
dot.edge('D', 'F', label='No')
dot.edge('F', 'G', label='Yes')
dot.edge('F', 'H', label='No')
dot.edge('H', 'I', label='Yes')
dot.edge('H', 'I', label='No', style='invis')  # Invisible edge to align nodes

# Additional formatting for clarity and aesthetics
dot.attr(label=r'\n\nTM Creation Decision Process Flowchart')
dot.attr(fontsize='12')

# Generate and display the cleaned-up flowchart
dot.render('output_files/TM_Creation_Decision_Process_Flowchart_Clean.dot', format='png', cleanup=True)
