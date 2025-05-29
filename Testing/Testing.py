from pymxs import runtime as rt
import os

#vr = rt.renderers.current
file_path = os.path("//HDI-FS/Technical/TECH/REVIT/Projects/HDI Systems/Lo-Post/23-12-15_Marketing Renders/vpost/Stainless Steel_v2.vfbl")
rt.vfbLayerMgr.loadLayersFromFile(file_path)

#rt.execute('vfbControl #loadglobalccpreset "Stainless Steel_v2.vfbl"')