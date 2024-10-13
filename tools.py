import logging

logging.basicConfig(level=logging.INFO)

def create_shape(shape_type, x, y, width, height, color="default"):
    """
    Creates a shape in Visio at given coordinates with specified dimensions.
    """
    # Adjust coordinates based on shape size to ensure it's fully within the canvas
    adjusted_x = max(width / 2, min(x, 100 - width / 2))
    adjusted_y = max(height / 2, min(y, 100 - height / 2))
    
    logging.info(f"Creating shape '{shape_type}' at ({adjusted_x}%, {adjusted_y}%) with dimensions {width}%x{height}% and color {color}")
    return f"Shape '{shape_type}' created at coordinates ({adjusted_x}%, {adjusted_y}%) with dimensions {width}%x{height}% and color {color}."

def connect_shapes(shape1, shape2):
    """
    Simulates connecting two shapes in Visio.
    """
    logging.info(f"Connecting shape '{shape1}' with shape '{shape2}'")
    return f"Shape '{shape1}' connected with shape '{shape2}'."

def modify_shape_properties(shape, property_name, value):
    """
    Simulates modifying properties of a shape in Visio.
    """
    logging.info(f"Modifying shape '{shape}': Setting {property_name} to {value}")
    return f"Shape '{shape}' updated: {property_name} set to {value}."