from flask import Flask, request, render_template
import win32com.client
import pythoncom

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])

def catia_parameters():
    if request.method == 'POST':
        a = to_float(request.form['a'])
        b = to_float(request.form['b'])
        c = to_float(request.form['c'])
        d = to_float(request.form['d'])
        percage = request.form['choicePercage']

        create_catia_part(a, b, c, d, percage)

        return 'Parameters sent to CATIA: a={}, b={}, c={}, d={}, percage={}'.format(a, b, c, d, percage)
    
    return render_template('form.html')

def to_float(value):
    try:
        return float(value.replace(',', '.'))
    except ValueError:
        return None

def create_catia_part(a, b, c, d, percage):

    pythoncom.CoInitialize()  # Needed if running in a thread or web application

    percage = bool(percage)

    try:
        # Connect to CATIA
        catia = win32com.client.Dispatch("CATIA.Application")
    except Exception as e:
        print(f"Error connecting to CATIA: {e}")
        return

    # Create a new part document
    documents = catia.Documents
    partDoc = documents.Add("Part")
    part = partDoc.Part
    bodies = part.Bodies
    body = bodies.Item(1)
    sketches = body.Sketches

    # Create a sketch on the xy plane
    xyPlane = part.OriginElements.PlaneXY
    sketch = sketches.Add(xyPlane)

    # Open sketch for drawing
    factory2D = sketch.OpenEdition()

    # Drawing a rectangle based on a and b parameters
    point1 = factory2D.CreatePoint(0, 0)
    point2 = factory2D.CreatePoint(a, 0)
    point3 = factory2D.CreatePoint(a, b)
    point4 = factory2D.CreatePoint(0, b)
    line1 = factory2D.CreateLine(0, 0, a, 0)
    line2 = factory2D.CreateLine(a, 0, a, b)
    line3 = factory2D.CreateLine(a, b, 0, b)
    line4 = factory2D.CreateLine(0, b, 0, 0)

    # Close the sketch edition
    sketch.CloseEdition()

    # Extrude the sketch
    shapeFactory = part.ShapeFactory
    pad = shapeFactory.AddNewPad(sketch, 20)  # Extrusion length

    # Update the part to apply changes
    part.Update()

if __name__ == '__main__':
    app.run(debug=True)