"""This module contains utility functions for the Batch Render extension."""

import uuid

from .config import Config

import pymxs

rt = pymxs.runtime


def generate_unique_id(obj):
    """Generates a new id for the object"""
    unique_ID = uuid.uuid4()
    rt.setUserProp(obj, Config.UUID_PARAMETER_NAME, unique_ID)
    rt.setSaveRequired(True)  # Makes 3ds max know that changes occurred to the file

    return unique_ID


def get_item_unique_id(obj):
    """Returns unique ID of an object.
    Creates a user-defined property (UDP) on the object if it does not already exist"""

    unique_ID = rt.getUserProp(obj, Config.UUID_PARAMETER_NAME)

    if not unique_ID:
        raise ValueError(f"Object does not have a unique ID: {obj}")

    return unique_ID


def get_item_by_id(unique_id):
    """Returns object by custom unique ID
    object_id is a custom property as 3ds max's built-in ID is not reliable"""

    for node in rt.objects:
        try:
            object_id = get_item_unique_id(node)
            if object_id == unique_id:
                return node
        except ValueError:
            pass

    raise ValueError(f'No object could not be found by unique ID "{unique_id}"')


def get_item_by_name(name):
    """Returns object by name"""

    node = rt.getNodeByName(name)
    if not node:
        raise ValueError(f"Object does not exist: {name}")

    return node
