import logging

logging.basicConfig(level=logging.INFO)

def create_shape(shape_type, x, y, width, height, color="default"):
    """
    Creates a shape in Visio at given coordinates with specified dimensions.
    """
    # Validate input types
    if not isinstance(x, (int, float)) or not isinstance(y, (int, float)):
        raise ValueError("Coordinates (x, y) must be numbers.")
    if not isinstance(width, (int, float)) or not isinstance(height, (int, float)):
        raise ValueError("Width and height must be numbers.")
    if not isinstance(color, str):
        raise ValueError("Color must be a string.")

    # Adjust coordinates based on shape size to ensure it's fully within the canvas
    adjusted_x = max(width / 2, min(x, 100 - width / 2))
    adjusted_y = max(height / 2, min(y, 100 - height / 2))
    
    logging.info(f"Creating shape '{shape_type}' at ({adjusted_x}%, {adjusted_y}%) with dimensions {width}%x{height}% and color {color}")
    
    return {
        "status": "success",
        "shape_type": shape_type,
        "position": {"x": adjusted_x, "y": adjusted_y},
        "dimensions": {"width": width, "height": height},
        "color": color
    }

def connect_shapes(shape1, shape2, connection_type="line"):
    """
    Connects two shapes in Visio with a specific connection type (line, arrow, etc.)
    """
    logging.info(f"Connecting shape '{shape1}' with shape '{shape2}' using {connection_type}.")
    return {
        "status": "success",
        "connection_type": connection_type,
        "shapes": [shape1, shape2]
    }

def modify_shape_properties(shape, property_name, value):
    """
    Simulates modifying properties of a shape in Visio.
    """
    valid_properties = ["color", "width", "height", "line_style"]  # Add more as needed

    if property_name not in valid_properties:
        logging.warning(f"Unknown property '{property_name}' for shape '{shape}'.")
        return {
            "status": "error",
            "message": f"Unknown property '{property_name}' for shape '{shape}'."
        }

    logging.info(f"Modifying shape '{shape}': Setting {property_name} to {value}")
    
    return {
        "status": "success",
        "shape": shape,
        "property": property_name,
        "new_value": value
    }
