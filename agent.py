import logging
from tools import create_shape, connect_shapes, modify_shape_properties

logging.basicConfig(level=logging.INFO)

class VisioAgent:
    def __init__(self):
        # Define supported commands
        self.commands = {
            "create_shape": self.create_shape,
            "connect_shapes": self.connect_shapes,
            "modify_properties": self.modify_properties,
        }

    def parse_command(self, command_text):
        """
        Use LangGraph's response to determine the appropriate Visio command.
        """
        logging.info(f"VisioAgent: Parsing command.")

        if "create a circle" in command_text.lower():
            return "create_shape", self.extract_shape_data(command_text)
        elif "connect" in command_text.lower():
            return "connect_shapes", self.extract_connect_data(command_text)
        elif "change color" in command_text.lower():
            return "modify_properties", self.extract_modify_data(command_text)
        else:
            logging.warning(f"Unsupported command: {command_text}")
            return None, None

    def execute_command(self, command_type, command_data):
        if command_type in self.commands:
            return self.commands[command_type](command_data)
        else:
            return {"error": f"Unsupported command '{command_type}'"}

    def create_shape(self, command_data):
        shape, x, y, color = command_data
        return create_shape(shape, x, y, color)

    def connect_shapes(self, command_data):
        shape1, shape2 = command_data
        return connect_shapes(shape1, shape2)

    def modify_properties(self, command_data):
        shape, property_name, value = command_data
        return modify_shape_properties(shape, property_name, value)

    def extract_shape_data(self, command_text):
        shape_type = "circle"
        x, y = 5.0, 5.0
        color = "blue"
        return shape_type, x, y, color

    def extract_connect_data(self, command_text):
        shape1, shape2 = "A", "B"
        return shape1, shape2

    def extract_modify_data(self, command_text):
        shape = "C"
        property_name = "color"
        value = "blue"
        return shape, property_name, value
