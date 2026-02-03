import turtle as tu
import re
import docx
import time

source = "Dhoni in Yellow Final"    # Mention the docx (Coordinate) file name without extension

try:
    data = docx.Document(f"{source}.docx")
except:
    print(f"Error: Could not find {source}.docx", flush=True)
    data = docx.Document()

pen = tu.Turtle()
screen = tu.Screen()
screen.setup(width=1.0, height=1.0)  # Fullscreen optimization

# Disable animation completely for maximum speed
tu.tracer(0, 0)
pen.hideturtle()
pen.speed(0)

# Flag to control instant completion
instant_complete = False

def complete_instantly():
    global instant_complete
    instant_complete = True
    print("Fast-forwarding to completion...", flush=True)

# Listen for Enter key
screen.listen()
screen.onkey(complete_instantly, "Return")

count = 0
total_shapes = len(data.paragraphs)
print(f"Found {total_shapes} shapes to draw.", flush=True)

start_time = time.time()

for i, content_paragraph in enumerate(data.paragraphs):
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
        
        # Move to first point
        x, y = map(float, coords[0])
        pen.goto(x, -y)
        
        pen.pendown()
        
        if draw == 1:
            pen.begin_fill()
            
        for j in range(1, len(coords)):
            x, y = map(float, coords[j])
            pen.goto(x, -y)
            
        # Close the shape
        x, y = map(float, coords[0])
        pen.goto(x, -y)
        
        if draw == 1:
            pen.end_fill()
    
    count += 1
    
    # Update screen every 20 shapes to create a fast animation effect
    if count % 20 == 0:
        screen.update()
        # print(f"Drawn {count}/{total_shapes} shapes...", flush=True)

# Final update
screen.update()
end_time = time.time()
print(f"Drawing finished in {end_time - start_time:.2f} seconds.", flush=True)

screen.mainloop()
