import array
import comtypes.client
from comtypes.client import GetBestInterface
import time
import math
import comtypes.automation
from comtypes.automation import VARIANT, VT_ARRAY, VT_DISPATCH
from ctypes import POINTER
import _ctypes  # <--- Required for catching the specific COMError

import time
import comtypes.client


class ConnectAutoCAD:
    def __init__(self):
        print("Attempting to connect to AutoCAD...")

        # Retry loop to bypass the "rejected by callee" error
        for attempt in range(5):
            try:
                self.acad = comtypes.client.GetActiveObject("AutoCAD.Application")
                print("   Success: Connected to AutoCAD!")

                # Setup your document and model space here
                self.doc = self.acad.ActiveDocument
                self.model_space = self.doc.ModelSpace
                break  # Exit the loop if successful

            except Exception as e:
                # If AutoCAD is busy, wait 2 seconds and try again
                if attempt < 4:
                    print(f"   AutoCAD is busy. Retrying in 2 seconds... (Attempt {attempt + 2}/5)")
                    time.sleep(2)
                else:
                    print("\nCRITICAL ERROR: Could not connect to AutoCAD.")
                    print("Make sure AutoCAD is open, a drawing is active, and no dialog boxes are open.")
                    raise e

    def _safe_com_call(self, action, max_retries=5, delay=0.1):
        """
        Wraps AutoCAD COM calls in a retry loop.
        If AutoCAD is busy (Rejected by callee), it waits a fraction of a second and tries again.
        """
        for attempt in range(max_retries):
            try:
                # Attempt to execute the AutoCAD command
                return action()
            except Exception as e:
                # Check if it is the specific "Rejected by callee" COM error
                if "-2147418111" in str(e) or "rejected by callee" in str(e).lower():
                    if attempt < max_retries - 1:
                        time.sleep(delay)  # Wait and try again
                        continue
                # If it's a different error or we are out of retries, raise the error
                raise e



    def true_color(self, rgb: tuple):
        major_version = self.acad.Version.split('.')[0]
        color_prog_id = f"AutoCAD.AcCmColor.{major_version}"

        try:
            # Request the TrueColor object from AutoCAD
            true_color = self.acad.GetInterfaceObject(color_prog_id)
        except Exception as e:
            print(f"   FAILED to load color interface ({color_prog_id}). Error: {e}")
            exit()

        # Define your RGB values (0 to 255)
        r, g, b = rgb[0], rgb[1], rgb[2]
        true_color.SetRGB(r, g, b)
        return true_color

    def apply_hatch(self, boundary_obj, hatch_pattern="SOLID", hatch_scale=1.0,
                    layer_name="NA", rgb=(255, 255, 255)):
        """
        Applies a hatch to a given boundary object.
        """
        if not boundary_obj:
            print("   [ERROR] Boundary object is missing.")
            return None

        # AutoCAD requires a very specific COM SAFEARRAY format for AppendOuterLoop
        def make_object_array(entities):
            SA_IDispatch = comtypes.automation._midlSAFEARRAY(POINTER(comtypes.automation.IDispatch))
            disp_pointers = [e.QueryInterface(comtypes.automation.IDispatch) for e in entities]
            return SA_IDispatch.from_param(disp_pointers)

        try:
            # 0 = acHatchPatternTypePreDefined
            hatch = self._safe_com_call(lambda: self.model_space.AddHatch(0, hatch_pattern, True))

            # Convert the boundary object into the specific COM array AutoCAD expects
            boundary_array = make_object_array([boundary_obj])
            hatch.AppendOuterLoop(boundary_array)

            hatch.PatternScale = float(hatch_scale)

            # Apply Color/Layer
            if layer_name.lower() == "na":
                hatch.TrueColor = self.true_color(rgb)
            else:
                hatch.Layer = layer_name
                hatch.Color = 256  # ByLayer

            # Evaluate and explicitly Update to render it in the viewport
            hatch.Evaluate()
            hatch.Update()

            print(f"   Success: Applied '{hatch_pattern}' hatch!")
            return hatch

        except Exception as e:
            print(f"   FAILED to apply hatch. Error: {e}")
            return None

    def ensure_layer(self, layer_name: str, rgb: tuple = (255, 255, 255)):
        """
        Checks if a layer exists.
        If it exists: Passes and returns the layer.
        If it does not exist: Creates it, applies a color, and returns the layer.
        """
        print(f"\nChecking for layer: '{layer_name}'...")

        try:
            # .Item() will throw an error if the layer does NOT exist
            layer = self.doc.Layers.Item(layer_name)
            print(f"   Result: Layer '{layer_name}' already exists. Passing.")
            return layer

        except Exception:
            # Since .Item() failed, we know the layer doesn't exist. Let's create it.
            print(f"   Result: Layer '{layer_name}' does not exist. Creating it now...")
            layer = self.doc.Layers.Add(layer_name)

            # Apply the color only because it is a newly created layer
            layer.TrueColor = self.true_color(rgb)
            print(f"   Success: New layer '{layer_name}' created with RGB{rgb}!")

            return layer

    def change_units(self, unit: int, linear_format: int, angular_format: int, linear_precision: int = 0,
                     angular_precision: int = 0):
        """
        Changes the drawing units and display formats for the active document.
        """
        try:
            self.doc.SetVariable("INSUNITS", unit)
            self.doc.SetVariable("LUNITS", linear_format)
            self.doc.SetVariable("LUPREC", linear_precision)
            self.doc.SetVariable("AUNITS", angular_format)
            self.doc.SetVariable("AUPREC", angular_precision)
            print("\nSuccess: Unit display types successfully updated!")
        except Exception as e:
            print(f"\nFAILED to update units. Error: {e}")

    def setup_dimstyle(self, style_name: str, arrow_size: float, arrow_style: str, text_size: float, text_gap: float,
                       text_position: int, precision: int, unit_format: int, suppress_ext_lines: bool, layer_name: str,
                       rgb: tuple):
        """
        Creates or updates a Dimension Style. Uses safe assignment to prevent COM crashes.
        """
        print(f"\nSetting up Dimension Style: '{style_name}'...")

        def safe_set_var(var_name, var_value):
            try:
                self.doc.SetVariable(var_name, var_value)
            except Exception:
                print(f"      [WARNING] AutoCAD rejected setting {var_name} to '{var_value}'. Skipping.")

        try:
            if arrow_style == "":
                arrow_style = "."

            safe_set_var("DIMASZ", float(arrow_size))
            safe_set_var("DIMBLK", str(arrow_style))
            safe_set_var("DIMLDRBLK", str(arrow_style))
            safe_set_var("DIMTXT", float(text_size))
            safe_set_var("DIMGAP", float(text_gap))
            safe_set_var("DIMTAD", int(text_position))
            safe_set_var("DIMLUNIT", int(unit_format))
            safe_set_var("DIMDEC", int(precision))
            safe_set_var("DIMDSEP", ".")

            if suppress_ext_lines:
                safe_set_var("DIMSE1", 1)
                safe_set_var("DIMSE2", 1)
            else:
                safe_set_var("DIMSE1", 0)
                safe_set_var("DIMSE2", 0)

            if layer_name.lower() == "na":
                tc = self.true_color(rgb)
                color_index = int(tc.ColorIndex)
                safe_set_var("DIMCLRD", color_index)
                safe_set_var("DIMCLRE", color_index)
                safe_set_var("DIMCLRT", color_index)
                safe_set_var("DIMLAYER", ".")
            else:
                safe_set_var("DIMCLRD", 256)
                safe_set_var("DIMCLRE", 256)
                safe_set_var("DIMCLRT", 256)
                safe_set_var("DIMLAYER", str(layer_name))

            try:
                dim_style = self.doc.DimStyles.Item(style_name)
                print(f"   Found existing DimStyle '{style_name}'. Updating settings...")
            except Exception:
                dim_style = self.doc.DimStyles.Add(style_name)
                print(f"   Created new DimStyle '{style_name}'.")

            dim_style.CopyFrom(self.doc)
            self.doc.ActiveDimStyle = dim_style
            print(f"   Success: '{style_name}' is now configured and set as active!")
            return dim_style

        except Exception as e:
            print(f"   FAILED to setup DimStyle. Error: {e}")
            return None

    def draw_line(self, xy1, xy2, layer_name, rgb, line_type, linetype_scale=1.0, global_width=0.0):
        print("\nDrawing the line...")

        try:
            # --- NEW: Global Width Logic ---
            # Standard Lines don't support Global Width.
            # If width > 0, we draw a 2-point LightWeightPolyline instead.
            if float(global_width) > 0.0:
                flat_array = array.array('d', [xy1[0], xy1[1], xy2[0], xy2[1]])
                # Uses the retry wrapper
                line_obj = self._safe_com_call(lambda: self.model_space.AddLightWeightPolyline(flat_array))
                # Set start and end width for vertex 0
                line_obj.SetWidth(0, float(global_width), float(global_width))
            else:
                start_point = array.array('d', [xy1[0], xy1[1], 0.0])
                end_point = array.array('d', [xy2[0], xy2[1], 0.0])
                # Uses the retry wrapper
                line_obj = self._safe_com_call(lambda: self.model_space.AddLine(start_point, end_point))

            # --- Color and Layer Setup ---
            if layer_name.lower() == "na":
                line_obj.TrueColor = self.true_color(rgb=rgb)
            else:
                line_obj.Layer = layer_name
                line_obj.Color = 256
                # Set default linetype to ByLayer only if they didn't specify a custom one
                if line_type.lower() == "na":
                    line_obj.Linetype = "ByLayer"
                line_obj.Lineweight = -1

            # --- Linetype Loading & Assignment ---
            if line_type.lower() not in ["na", "bylayer", "byblock", "continuous"]:
                try:
                    # 1. Attempt to apply the linetype directly (works if already loaded)
                    line_obj.Linetype = line_type
                except Exception:
                    # 2. If it fails, attempt to load it on the fly
                    print(f"   [INFO] Linetype '{line_type}' not found in document. Attempting to load...")
                    try:
                        self.doc.Linetypes.Load(line_type, "acadiso.lin")
                        line_obj.Linetype = line_type
                        print(f"   Success: Loaded '{line_type}' from acadiso.lin.")
                    except Exception:
                        try:
                            self.doc.Linetypes.Load(line_type, "acad.lin")
                            line_obj.Linetype = line_type
                            print(f"   Success: Loaded '{line_type}' from acad.lin.")
                        except Exception as e:
                            print(f"   [WARNING] FAILED to load linetype '{line_type}'. Error: {e}")

                # 3. Apply the Linetype Scale
                try:
                    line_obj.LinetypeScale = float(linetype_scale)
                except Exception as e:
                    print(f"   [WARNING] Could not apply linetype scale. Error: {e}")

            # self.acad.Application.ZoomExtents()
            print(f"   Success: Line drawn from {list(xy1)} to {list(xy2)}!")

            return line_obj

        except Exception as e:
            print(f"   FAILED to draw line. Error: {e}")
            return None

    def draw_polyline(self, points, layer_name="NA", rgb=(255, 255, 255), bulges=None):
        """
        Draws a LightWeightPolyline.
        Takes a flat list of coordinates: [x1, y1, x2, y2, ...]
        Accepts an optional 'bulges' dictionary to turn straight segments into arcs.
        """
        print("\nDrawing Polyline...")
        flat_array = array.array('d', points)

        try:
            # Uses the retry wrapper
            polyline = self._safe_com_call(lambda: self.doc.ModelSpace.AddLightWeightPolyline(flat_array))

            # Apply properties
            if layer_name.lower() == "na":
                polyline.TrueColor = self.true_color(rgb=rgb)
            else:
                polyline.Layer = layer_name
                polyline.Color = 256
                polyline.Linetype = "ByLayer"
                polyline.Lineweight = -1

            # --- Apply Arcs (Bulges) to specific segments ---
            if bulges:
                for vertex_index, bulge_value in bulges.items():
                    try:
                        # SetBulge takes the vertex index and the bulge float value
                        polyline.SetBulge(vertex_index, float(bulge_value))
                    except Exception as e:
                        print(f"   [WARNING] Failed to apply bulge at vertex {vertex_index}. Error: {e}")

            # self.acad.Application.ZoomExtents()
            print("   Success: Polyline drawn!")
            return polyline

        except Exception as e:
            print(f"   FAILED to draw polyline. Error: {e}")
            return None

    def draw_arc_2p(self, xy1, xy2, radius: float, layer_name="NA", rgb=(255, 255, 255), line_type="NA",
                    linetype_scale=1.0, global_width=0.0):
        """
        Draws an arc from a Start point (xy1) to an End point (xy2) with a specific radius.
        Draws counter-clockwise. Automatically calculates the center point and angles.
        """
        print(f"\nDrawing arc from {xy1} to {xy2} with radius {radius}...")

        x1, y1 = xy1[0], xy1[1]
        x2, y2 = xy2[0], xy2[1]

        # 1. Calculate the distance between the two points
        dx = x2 - x1
        dy = y2 - y1
        d = math.hypot(dx, dy)

        # Geometry Check: If points are further apart than the diameter, an arc is mathematically impossible.
        if d > 2 * radius:
            print(f"   [ERROR] Points are too far apart ({d:.2f}) for a radius of {radius}.")
            return None

        # 2. Calculate the Midpoint between the two coordinates
        mx = x1 + dx / 2.0
        my = y1 + dy / 2.0

        # 3. Calculate distance from midpoint to the center of the circle (Pythagorean theorem)
        # Using max(0, val) prevents floating-point crash if d is exactly equal to 2*radius
        h_squared = (radius ** 2) - ((d / 2.0) ** 2)
        h = math.sqrt(max(0, h_squared))

        # 4. Calculate the center point
        # To draw counter-clockwise from xy1 to xy2, we offset perpendicular to the left
        nx = -dy / d
        ny = dx / d

        cx = mx + h * nx
        cy = my + h * ny

        # 5. Calculate the Start and End angles in radians for AutoCAD
        start_rad = math.atan2(y1 - cy, x1 - cx)
        end_rad = math.atan2(y2 - cy, x2 - cx)

        try:
            # --- Global Width Logic ---
            if float(global_width) > 0.0:
                # If width is requested, we draw it as a thickened Polyline arc
                flat_array = array.array('d', [x1, y1, x2, y2])

                # Uses the retry wrapper
                arc_obj = self._safe_com_call(lambda: self.model_space.AddLightWeightPolyline(flat_array))

                # Calculate the Bulge for the polyline (tan(delta_angle / 4))
                delta_angle = end_rad - start_rad
                if delta_angle < 0:
                    delta_angle += 2 * math.pi

                bulge = math.tan(delta_angle / 4.0)

                arc_obj.SetBulge(0, bulge)
                arc_obj.SetWidth(0, float(global_width), float(global_width))

            else:
                # Standard AcDbArc creation if width is 0
                center_point = array.array('d', [cx, cy, 0.0])
                arc_obj = self._safe_com_call(lambda: self.model_space.AddArc(center_point, radius, start_rad, end_rad))

            # --- Color and Layer Setup ---
            if layer_name.lower() == "na":
                arc_obj.TrueColor = self.true_color(rgb=rgb)
            else:
                arc_obj.Layer = layer_name
                arc_obj.Color = 256
                if line_type.lower() == "na":
                    arc_obj.Linetype = "ByLayer"
                arc_obj.Lineweight = -1

            # --- Linetype Loading & Assignment ---
            if line_type.lower() not in ["na", "bylayer", "byblock", "continuous"]:
                try:
                    arc_obj.Linetype = line_type
                except Exception:
                    print(f"   [INFO] Linetype '{line_type}' not found. Attempting to load...")
                    try:
                        self.doc.Linetypes.Load(line_type, "acadiso.lin")
                        arc_obj.Linetype = line_type
                        print(f"   Success: Loaded '{line_type}'.")
                    except Exception:
                        try:
                            self.doc.Linetypes.Load(line_type, "acad.lin")
                            arc_obj.Linetype = line_type
                            print(f"   Success: Loaded '{line_type}'.")
                        except Exception as e:
                            print(f"   [WARNING] FAILED to load linetype '{line_type}'. Error: {e}")

                try:
                    arc_obj.LinetypeScale = float(linetype_scale)
                except Exception as e:
                    print(f"   [WARNING] Could not apply scale. Error: {e}")

            # self.acad.Application.ZoomExtents()
            print("   Success: Arc drawn!")
            return arc_obj

        except Exception as e:
            print(f"   FAILED to draw arc. Error: {e}")
            return None

    def join_objects(self, objects_list, global_width=0.0):
        """
        Joins separate lines and arcs into a single, seamless Polyline.
        Applies a global width to the final unified shape so corners are perfectly smooth.
        """
        print("\nJoining objects for smooth corners...")
        if not objects_list or len(objects_list) < 2:
            print("   [WARNING] Need at least 2 objects to join.")
            return

        try:
            # 1. Join everything. The JOIN command merges them into a single Polyline.
            primary_handle = objects_list[0].Handle

            # Construct the LISP JOIN command
            keystrokes = '\x03\x03(command "_.JOIN" '
            for obj in objects_list:
                if obj:
                    keystrokes += f'(handent "{obj.Handle}") '
            keystrokes += ' "")\n'

            # --- FIX: Wrapped in the retry block to prevent "Rejected by Callee" ---
            self._safe_com_call(lambda: self.doc.SendCommand(keystrokes))

            # Give AutoCAD a split second to physically update its database
            time.sleep(0.2)

            # 2. Apply the Global Width
            if float(global_width) > 0.0:
                pedit_keystrokes = f'\x03\x03(command "_.PEDIT" (handent "{primary_handle}") "_Width" {global_width} "")\n'
                self._safe_com_call(lambda: self.doc.SendCommand(pedit_keystrokes))

            # self.acad.Application.ZoomExtents()
            print("   Success: Objects joined perfectly!")

        except Exception as e:
            print(f"   FAILED to join objects. Error: {e}")

    def fillet_objects(self, obj1, obj2, radius: float):
        """
        Creates a fillet between two lines using AutoLISP.
        Includes geometric error trapping for 'Invalid input' COM errors.
        """
        print(f"\nApplying fillet with radius {radius}...")

        if not obj1 or not obj2:
            print("   [ERROR] Missing objects. Cannot apply fillet.")
            return

        try:
            # 1. Safely set the fillet radius
            self.doc.SetVariable("FILLETRAD", float(radius))

            # 2. Construct the safest single-line LISP injection
            # \x03\x03 clears the command line.
            # (command "_.FILLET" ...) executes the whole thing as one LISP transaction.
            lisp_command = f'\x03\x03(command "_.FILLET" (handent "{obj1.Handle}") (handent "{obj2.Handle}"))\n'

            # 3. Inject it
            self.doc.SendCommand(lisp_command)

            print("   Success: Fillet command executed!")

        except Exception as e:
            print(f"   FAILED to apply fillet. AutoCAD rejected the command.")
            print(f"   Error: {e}")

            # Diagnostic help for the specific error you received
            if "-2145386493" in str(e) or "Invalid input" in str(e):
                print("\n   [ GEOMETRY REJECTION DETECTED ]")
                print("   AutoCAD aborted the fillet. Check the following:")
                print("   1. RADIUS TOO LARGE: Are your line segments physically longer than 28.38?")
                print("   2. Z-AXIS MISMATCH: Ensure both lines have a Z-coordinate of exactly 0.0.")
                print(
                    "   3. PARALLEL LINES: If the lines are perfectly parallel, they must overlap vertically/horizontally.")
                print(
                    "   -> TEST: Try running this exact function with radius=0.0. If it works, your lines are fine, but 28.38 is too big for the space.")

    def draw_dimension(self, xy1, xy2, text_xy, dim_type="ALIGNED", text_override="",
                       ext_offset=0.0, ext_extend=0.0, fixed_ext_len=0.0,
                       layer_name="NA", rgb=(255, 255, 255)):
        """
        Draws a dimension based on the requested type.
        dim_type options:
          - "ALIGNED" (Default: Diagonal / True Length)
          - "H" (Horizontal / Linear)
          - "V" (Vertical / Linear)
        """
        dim_t = str(dim_type).strip().upper()
        print(f"\nDrawing {dim_t} dimension from {xy1} to {xy2}...")

        pt1 = array.array('d', [float(xy1[0]), float(xy1[1]), 0.0])
        pt2 = array.array('d', [float(xy2[0]), float(xy2[1]), 0.0])
        text_pt = array.array('d', [float(text_xy[0]), float(text_xy[1]), 0.0])

        # --- 1. SAFE STRING CONVERSION ---
        raw_fixed = str(fixed_ext_len).strip()
        safe_fixed = float(raw_fixed) if raw_fixed != "" else 0.0

        raw_offset = str(ext_offset).strip()
        safe_offset = float(raw_offset) if raw_offset != "" else 0.0

        raw_extend = str(ext_extend).strip()
        safe_extend = float(raw_extend) if raw_extend != "" else 0.0

        # --- 2. CALCULATE MATH OFFSET FOR FIXED LINES ---
        if safe_fixed > 0.0:
            if dim_t == "H":
                # For horizontal dims, check Y distance
                total_distance = abs(float(text_xy[1]) - float(xy1[1]))
            elif dim_t == "V":
                # For vertical dims, check X distance
                total_distance = abs(float(text_xy[0]) - float(xy1[0]))
            else:
                # ALIGNED: Calculate perpendicular distance
                x1, y1 = float(xy1[0]), float(xy1[1])
                x2, y2 = float(xy2[0]), float(xy2[1])
                x0, y0 = float(text_xy[0]), float(text_xy[1])
                try:
                    num = abs((y2 - y1) * x0 - (x2 - x1) * y0 + x2 * y1 - y2 * x1)
                    den = math.hypot(y2 - y1, x2 - x1)
                    total_distance = num / den
                except ZeroDivisionError:
                    total_distance = 0.0

            calculated_offset = total_distance - safe_fixed
            safe_offset = max(calculated_offset, 0.0)

        try:
            # --- 3. DRAW THE DIMENSION ---
            if dim_t == "H":
                dim_obj = self._safe_com_call(lambda: self.model_space.AddDimRotated(pt1, pt2, text_pt, 0.0))
            elif dim_t == "V":
                dim_obj = self._safe_com_call(lambda: self.model_space.AddDimRotated(pt1, pt2, text_pt, math.pi / 2.0))
            else:
                dim_obj = self._safe_com_call(lambda: self.model_space.AddDimAligned(pt1, pt2, text_pt))

            # --- Text Override ---
            if text_override:
                dim_obj.TextOverride = str(text_override)

            # --- Extension Line Controls ---
            dim_obj.ExtLine1Suppress = False
            dim_obj.ExtLine2Suppress = False

            if safe_offset > 0.0:
                dim_obj.ExtensionLineOffset = safe_offset
            if safe_extend > 0.0:
                dim_obj.ExtensionLineExtend = safe_extend

            # --- Color and Layer Setup ---
            if layer_name.lower() == "na":
                dim_obj.TrueColor = self.true_color(rgb=rgb)
            else:
                dim_obj.Layer = layer_name
                dim_obj.Color = 256  # ByLayer

            print(f"   Success: {dim_t} Dimension drawn successfully!")
            return dim_obj

        except Exception as e:
            print(f"   FAILED to draw dimension. Error: {e}")
            return None

    import math
    import array

    def draw_text(self, text_string, xy, height=100.0, alignment="LEFT",
                  orientation="HORIZONTAL", underline=False, bold=False,
                  multi_line=False, mtext_width=0.0,
                  layer_name="NA", rgb=(255, 255, 255)):
        """
        Draws Text in AutoCAD.
        Supports Underline, Bold (Bold requires multi_line=True),
        and Orientation (HORIZONTAL, VERTICAL_LEFT, VERTICAL_RIGHT).
        """
        print(f"\nDrawing text at {xy}...")

        # --- 1. SAFE CONVERSIONS ---
        raw_height = str(height).strip()
        safe_height = float(raw_height) if raw_height != "" else 100.0

        raw_width = str(mtext_width).strip()
        safe_width = float(raw_width) if raw_width != "" else 0.0

        is_underlined = str(underline).strip().upper() in ["TRUE", "1", "YES"] or underline is True
        is_bold = str(bold).strip().upper() in ["TRUE", "1", "YES"] or bold is True
        is_multi = str(multi_line).strip().upper() in ["TRUE", "1", "YES"] or multi_line is True

        raw_text = str(text_string)

        # --- 2. ORIENTATION / ROTATION LOGIC ---
        orient_str = str(orientation).strip().upper()
        if orient_str in ["VERTICAL_LEFT", "HEAD_LEFT"]:
            # Rotated 90 degrees (reads bottom-to-top)
            rotation_angle = math.pi / 2.0
        elif orient_str in ["VERTICAL_RIGHT", "HEAD_RIGHT"]:
            # Rotated -90 degrees (reads top-to-bottom)
            rotation_angle = -math.pi / 2.0
        else:
            # Default horizontal
            rotation_angle = 0.0

        # --- 3. FORMATTING CODES ---
        if is_multi:
            display_text = raw_text

            if is_underlined:
                display_text = "\\L" + display_text

            if is_bold:
                # \\f turns on font formatting. Arial is a TrueType font. b1 means Bold ON.
                display_text = "\\fArial|b1;" + display_text
        else:
            display_text = raw_text

            if is_bold:
                print("   WARNING: AutoCAD Single Line Text does not support Bold. Skipping bold formatting.")

            if is_underlined:
                display_text = "%%u" + display_text

        insert_pt = array.array('d', [float(xy[0]), float(xy[1]), 0.0])
        align_str = str(alignment).strip().upper()

        try:
            if is_multi:
                # ==========================================
                # MULTI-LINE TEXT (MText) LOGIC
                # ==========================================
                text_obj = self._safe_com_call(lambda: self.model_space.AddMText(insert_pt, safe_width, display_text))
                text_obj.Height = safe_height

                # MText uses AttachmentPoint for alignment
                if align_str == "CENTER":
                    text_obj.AttachmentPoint = 8  # BottomCenter
                elif align_str == "RIGHT":
                    text_obj.AttachmentPoint = 9  # BottomRight
                elif align_str in ["MIDDLE", "MIDDLE_CENTER"]:
                    text_obj.AttachmentPoint = 5  # Dead Center
                else:
                    text_obj.AttachmentPoint = 7  # BottomLeft

                text_obj.InsertionPoint = insert_pt

                # Apply Rotation
                text_obj.Rotation = rotation_angle

            else:
                # ==========================================
                # SINGLE LINE TEXT LOGIC
                # ==========================================
                text_obj = self._safe_com_call(lambda: self.model_space.AddText(display_text, insert_pt, safe_height))

                # Single Line uses Alignment and TextAlignmentPoint
                if align_str != "LEFT":
                    if align_str == "CENTER":
                        text_obj.Alignment = 1
                    elif align_str == "RIGHT":
                        text_obj.Alignment = 2
                    elif align_str in ["MIDDLE", "MIDDLE_CENTER"]:
                        text_obj.Alignment = 4

                    text_obj.TextAlignmentPoint = insert_pt

                # Apply Rotation
                text_obj.Rotation = rotation_angle

            # --- LAYER AND COLOR SETUP ---
            if layer_name.lower() == "na":
                text_obj.TrueColor = self.true_color(rgb=rgb)
            else:
                text_obj.Layer = layer_name
                text_obj.Color = 256

            print(f"   Success: {'Multi-Line' if is_multi else 'Single Line'} Text drawn successfully!")
            return text_obj

        except Exception as e:
            print(f"   FAILED to draw text. Error: {e}")
            return None

    def draw_leader(self, points, text_string="", head_type="CLOSED", misc_type="LINE",
                    arrow_size=50.0, text_height=100.0,
                    direction="RIGHT", justify="LEFT",
                    layer_name="NA", rgb=(255, 255, 255)):
        """
        Draws a Leader line.
        Safely enforces Line vs Spline and is Loop-Safe for fast iterations.
        """
        import time  # Required for the loop-safe delay

        print(f"\nDrawing leader with {len(points)} points...")

        flat_pts = []
        for pt in points:
            flat_pts.extend([float(pt[0]), float(pt[1]), 0.0])
        pts_array = array.array('d', flat_pts)

        raw_arrow = str(arrow_size).strip()
        safe_arrow = float(raw_arrow) if raw_arrow != "" else 50.0

        raw_height = str(text_height).strip()
        safe_height = float(raw_height) if raw_height != "" else 100.0

        try:
            # --- 1. APPLY INTERNAL TEXT JUSTIFICATION ---
            just_str = str(justify).strip().upper()
            if just_str in ["CENTER", "MIDDLE", "MIDDLE_CENTER"]:
                format_code = "\\pxqc;"
            elif just_str == "RIGHT":
                format_code = "\\pxqr;"
            else:
                format_code = "\\pxql;"

            raw_text = str(text_string) if str(text_string).strip() != "" else " "
            display_text = format_code + raw_text

            text_pt = array.array('d', [float(points[-1][0]), float(points[-1][1]), 0.0])

            mtext_obj = self._safe_com_call(lambda: self.model_space.AddMText(text_pt, 0.0, display_text))
            mtext_obj.Height = safe_height

            # --- 2. APPLY LEADER DIRECTION (Attachment Point) ---
            dir_str = str(direction).strip().upper()
            if dir_str == "LEFT":
                mtext_obj.AttachmentPoint = 6
            elif dir_str in ["MIDDLE", "MIDDLE_CENTER", "CENTER_MIDDLE"]:
                mtext_obj.AttachmentPoint = 5
            elif dir_str == "CENTER":
                mtext_obj.AttachmentPoint = 8
            else:
                mtext_obj.AttachmentPoint = 4

            mtext_obj.InsertionPoint = text_pt

            # --- 3. DETERMINE MISC TYPE ---
            misc_dict = {
                "LINE_NO_ARROW": 0, "LINE": 2,
                "SPLINE_NO_ARROW": 1, "SPLINE": 3
            }
            m_type_str = str(misc_type).strip().upper()
            leader_code = misc_dict.get(m_type_str, 1)

            # --- 4. CREATE THE LEADER ---
            leader_obj = self._safe_com_call(lambda: self.model_space.AddLeader(pts_array, mtext_obj, leader_code))

            if leader_obj.Type != leader_code:
                leader_obj.Type = leader_code

            # --- 5. APPLY ARROWHEAD TYPE ---
            arrow_dict = {
                "CLOSED": 0, "BLANK": 1, "DOT": 3, "TICK": 4,
                "OBLIQUE": 5, "OPEN": 6, "OPEN90": 9, "OPEN30": 10,
                "DOT_SMALL": 11, "EMPTY_DOT": 12, "DOT_BLANK": 12,
                "BOX": 15, "NONE": 19
            }
            h_type_str = str(head_type).strip().upper()
            leader_obj.ArrowheadType = arrow_dict.get(h_type_str, 0)
            leader_obj.ArrowheadSize = safe_arrow

            # --- 6. LAYER AND COLOR SETUP ---
            if layer_name.lower() == "na":
                leader_obj.TrueColor = self.true_color(rgb=rgb)
                mtext_obj.TrueColor = self.true_color(rgb=rgb)
            else:
                leader_obj.Layer = layer_name
                leader_obj.Color = 256
                mtext_obj.Layer = layer_name
                mtext_obj.Color = 256

            # --- 7. SAFE SCREEN REFRESH ---
            # Isolated in try/excepts so if AutoCAD is laggy, it won't crash the script
            try:
                leader_obj.Update()
            except:
                pass

            print("   Success: Leader drawn successfully!")

            # Tiny delay to prevent COM race conditions in tight loops
            time.sleep(0.05)

            return leader_obj, mtext_obj

        except Exception as e:
            print(f"   FAILED to draw leader. Error: {e}")
            return None, None

    def draw_triangle(self, A, B, C,
                      hatch_pattern="SOLID",
                      hatch_scale=1.0,
                      layer_name="NA",
                      rgb=(255, 255, 255)):

        pts = array.array('d', [
            float(A[0]), float(A[1]),
            float(B[0]), float(B[1]),
            float(C[0]), float(C[1]),
            float(A[0]), float(A[1])
        ])

        tri = self._safe_com_call(lambda: self.model_space.AddLightWeightPolyline(pts))
        tri.Closed = True

        if layer_name.lower() == "na":
            tri.TrueColor = self.true_color(rgb=rgb)
        else:
            tri.Layer = layer_name
            tri.Color = 256

        def p(xy):
            return array.array('d', [float(xy[0]), float(xy[1]), 0.0])


        return tri

    def draw_circle(self, xy, radius: float, layer_name="NA", rgb=(255, 255, 255),
                    line_type="NA", linetype_scale=1.0, global_width=0.0):
        """
        Draws a circle at a specified center point with a given radius.
        Supports global width by dynamically converting to a closed Polyline if width > 0.
        """
        print(f"\nDrawing circle at {xy} with radius {radius}...")

        try:
            # --- Global Width Logic ---
            if float(global_width) > 0.0:
                # Standard Circles don't support Global Width.
                # If width > 0, we draw a 2-point LightWeightPolyline with two semicircles.
                x, y = float(xy[0]), float(xy[1])
                r = float(radius)

                # Create points for the left and right sides of the circle
                flat_array = array.array('d', [x - r, y, x + r, y])

                # Uses the retry wrapper
                circle_obj = self._safe_com_call(lambda: self.model_space.AddLightWeightPolyline(flat_array))

                # A bulge of 1.0 or -1.0 creates a perfect semicircle
                circle_obj.SetBulge(0, 1.0)
                circle_obj.SetBulge(1, 1.0)
                circle_obj.Closed = True

                # Apply the width
                circle_obj.SetWidth(0, float(global_width), float(global_width))
                circle_obj.SetWidth(1, float(global_width), float(global_width))
            else:
                # Standard AcDbCircle creation if width is 0
                center_point = array.array('d', [float(xy[0]), float(xy[1]), 0.0])
                circle_obj = self._safe_com_call(lambda: self.model_space.AddCircle(center_point, float(radius)))

            # --- Color and Layer Setup ---
            if layer_name.lower() == "na":
                circle_obj.TrueColor = self.true_color(rgb=rgb)
            else:
                circle_obj.Layer = layer_name
                circle_obj.Color = 256
                # Set default linetype to ByLayer only if they didn't specify a custom one
                if line_type.lower() == "na":
                    circle_obj.Linetype = "ByLayer"
                circle_obj.Lineweight = -1

            # --- Linetype Loading & Assignment ---
            if line_type.lower() not in ["na", "bylayer", "byblock", "continuous"]:
                try:
                    # 1. Attempt to apply the linetype directly
                    circle_obj.Linetype = line_type
                except Exception:
                    # 2. If it fails, attempt to load it on the fly
                    print(f"   [INFO] Linetype '{line_type}' not found in document. Attempting to load...")
                    try:
                        self.doc.Linetypes.Load(line_type, "acadiso.lin")
                        circle_obj.Linetype = line_type
                        print(f"   Success: Loaded '{line_type}' from acadiso.lin.")
                    except Exception:
                        try:
                            self.doc.Linetypes.Load(line_type, "acad.lin")
                            circle_obj.Linetype = line_type
                            print(f"   Success: Loaded '{line_type}' from acad.lin.")
                        except Exception as e:
                            print(f"   [WARNING] FAILED to load linetype '{line_type}'. Error: {e}")

                # 3. Apply the Linetype Scale
                try:
                    circle_obj.LinetypeScale = float(linetype_scale)
                except Exception as e:
                    print(f"   [WARNING] Could not apply linetype scale. Error: {e}")

            print("   Success: Circle drawn!")
            return circle_obj

        except Exception as e:
            print(f"   FAILED to draw circle. Error: {e}")
            return None
def set_ucs_by_click(app_instance):
        print("Go to AutoCAD and click a point on the screen...")

        try:
            # 1. Use a lambda function to wrap the GetPoint command safely
            clicked_point = app_instance._safe_com_call(
                lambda: app_instance.doc.Utility.GetPoint(None, "\nClick to set new UCS origin: ")
            )

            # 2. Setup the coordinate arrays for comtypes
            origin = array.array('d', clicked_point)
            x_axis = array.array('d', [clicked_point[0] + 1.0, clicked_point[1], clicked_point[2]])
            y_axis = array.array('d', [clicked_point[0], clicked_point[1] + 1.0, clicked_point[2]])

            # 3. Add the UCS safely using lambda
            new_ucs = app_instance._safe_com_call(
                lambda: app_instance.doc.UserCoordinateSystems.Add(origin, x_axis, y_axis, "ClickUCS")
            )

            # 4. To set a property (like ActiveUCS), lambda doesn't work well,
            # so we define a quick inline function to pass to _safe_com_call
            def apply_active_ucs():
                app_instance.doc.ActiveUCS = new_ucs

            app_instance._safe_com_call(apply_active_ucs)

            print("Success! UCS updated to your clicked location.")

        except Exception as e:
            print(f"Action cancelled or failed. Details: {e}")

