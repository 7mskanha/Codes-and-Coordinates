import turtle as tu
import re
import docx

# Source file for coordinates
source = "World cup final shot"

try:
    data = docx.Document(f"{source}.docx")
except:
    print(f"Error: Could not find {source}.docx")
    data = docx.Document()

pen = tu.Turtle()
screen = tu.Screen()

# Setup drawing speed
tu.tracer(10)  # Update screen every 10th frame for faster drawing
tu.hideturtle()
pen.speed(0)
pen.hideturtle()

for content_paragraph in data.paragraphs:
    content = content_paragraph.text
    
    # Check for outline flag
    is_outline = False
    if "outline" in content:
        is_outline = True
        content = content.replace(" outline", "")
    
    # Extract color
    color_match = re.search(r'\(([\d\.]+), ([\d\.]+), ([\d\.]+)\)', content)
    if color_match:
        r = float(color_match.group(1))
        g = float(color_match.group(2))
        b = float(color_match.group(3))
        pen.color(r, g, b)
        pen.fillcolor(r, g, b)
        
        # If outline, set pen to black and don't fill
        if is_outline:
            pen.pencolor(0, 0, 0)
            pen.pensize(1)
            draw = 0
        else:
            pen.pensize(1)
            draw = 1
            
    # Extract coordinates
    coords = re.findall(r'\(([\d\.-]+), ([\d\.-]+)\)', content)
    
    if coords:
        pen.penup()
        x, y = map(float, coords[0])
        pen.goto(x, -y)
        pen.pendown()
        
        if draw == 1:
            pen.begin_fill()
            
        for i in range(1, len(coords)):
            x, y = map(float, coords[i])
            pen.goto(x, -y)
            
        # Close the shape
        x, y = map(float, coords[0])
        pen.goto(x, -y)
        
        if draw == 1:
            pen.end_fill()

screen.mainloop()
