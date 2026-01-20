"""Main entry for 3ds max V-Ray batch render program"""

import argparse
import os
import sys

from mocks import pymxs

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--debug", action="store_true", help="Run in debug mode (outside of 3ds Max)"
    )
    parser.add_argument("--pipe", type=str, help="Named pipe for communication")
    args = parser.parse_args()

    if args.debug:
        print("Debug mode enabled")

        sys.modules["pymxs"] = pymxs

# Add the project root to Python path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)
os.chdir(project_root)

import batch_render
import log_window

# Source - https://stackoverflow.com/a/73456872
# Posted by logworthy, modified by community. See post 'Timeline' for change history
# Retrieved 2026-01-13, License - CC BY-SA 4.0

import importlib
import os
import types
import pathlib


def get_package_dependencies(package) -> tuple:
    """Recursively traverses the package dependencies
    Args:
        package: Package to traverse
    Returns:
        tuple: Tuple of node package dictionary and node depth dictionary
    """
    assert hasattr(package, "__package__")
    fn = package.__file__
    fn_dir = os.path.dirname(fn) + os.sep
    node_set = {fn}  # set of module filenames
    node_depth_dict = {fn: 0}  # tracks the greatest depth that we've seen for each node
    node_pkg_dict = {fn: package}  # mapping of module filenames to module objects
    link_set = set()  # tuple of (parent module filename, child module filename)
    del fn

    def dependency_traversal_recursive(module, depth):
        """Recursively traverses the package dependencies
        Args:
            module: Module to traverse
            depth: Depth of the module
        """
        for module_child in vars(module).values():

            # skip anything that isn't a module
            if not isinstance(module_child, types.ModuleType):
                continue

            fn_child = getattr(module_child, "__file__", None)

            # skip anything without a filename or outside the package
            if (fn_child is None) or (not fn_child.startswith(fn_dir)):
                continue

            # have we seen this module before? if not, add it to the database
            if not fn_child in node_set:
                node_set.add(fn_child)
                node_depth_dict[fn_child] = depth
                node_pkg_dict[fn_child] = module_child

            # set the depth to be the deepest depth we've encountered the node
            node_depth_dict[fn_child] = max(depth, node_depth_dict[fn_child])

            # have we visited this child module from this parent module before?
            if not ((module.__file__, fn_child) in link_set):
                link_set.add((module.__file__, fn_child))
                dependency_traversal_recursive(module_child, depth + 1)
            else:
                raise ValueError("Cycle detected in dependency graph!")

    dependency_traversal_recursive(package, 1)
    return node_pkg_dict, node_depth_dict


def reload(module) -> None:
    """Reloads the module and its dependencies
    Args:
        module: Module to reload
    """
    node_pkg_dict, node_depth_dict = get_package_dependencies(module)
    for d, v in sorted([(d, v) for v, d in node_depth_dict.items()], reverse=True):
        print("Reloading %s" % pathlib.Path(v).name)
        importlib.reload(node_pkg_dict[v])


reload(log_window)
reload(batch_render)


def main():
    """Main function"""
    from qtmax import GetQMaxMainWindow

    try:
        main_window = batch_render.BatchRenderDialog(parent=GetQMaxMainWindow())
        main_window.show()
    except Exception as main_error:
        log_window.console.close()
        print(f"Error in main: {str(main_error)}")
        import traceback

        traceback.print_exc()
        raise main_error
    finally:
        sys.exit()


if __name__ == "__main__":
    main()
