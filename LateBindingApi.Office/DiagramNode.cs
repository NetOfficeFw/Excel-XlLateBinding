//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class DiagramNode : _IMsoDispObj
	{
		#region Construction

		public DiagramNode(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public DiagramNode(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public DiagramNode(COMObject replacedObject) : base(replacedObject)
		{
		}

		public DiagramNode()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.DiagramNodeChildren Children
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Children");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.DiagramNodeChildren newClass = new LateBindingApi.Office.DiagramNodeChildren(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.Shape Shape
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Shape");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.Shape newClass = new LateBindingApi.Office.Shape(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.DiagramNode Root
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Root");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.DiagramNode newClass = new LateBindingApi.Office.DiagramNode(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.IMsoDiagram Diagram
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Diagram");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IMsoDiagram newClass = new LateBindingApi.Office.IMsoDiagram(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoOrgChartLayoutType Layout
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Layout");
				return (LateBindingApi.Office.Enums.MsoOrgChartLayoutType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Layout", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.Shape TextShape
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextShape");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.Shape newClass = new LateBindingApi.Office.Shape(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.DiagramNode AddNode(LateBindingApi.Office.Enums.MsoRelativeNodePosition pos, LateBindingApi.Office.Enums.MsoDiagramNodeType nodeType)
		{
			object[] paramArray = new object[2];
			paramArray[0] = pos;
			paramArray[1] = nodeType;
			object returnValue = Invoker.MethodReturn(this, "AddNode", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.DiagramNode newClass = new LateBindingApi.Office.DiagramNode(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void MoveNode(LateBindingApi.Office.DiagramNode targetNode, LateBindingApi.Office.Enums.MsoRelativeNodePosition pos)
		{
			object[] paramArray = new object[2];
			paramArray.SetValue(targetNode,0);
			paramArray[1] = pos;
			Invoker.Method(this, "MoveNode", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void ReplaceNode(LateBindingApi.Office.DiagramNode targetNode)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(targetNode,0);
			Invoker.Method(this, "ReplaceNode", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void SwapNode(LateBindingApi.Office.DiagramNode targetNode, bool swapChildren)
		{
			object[] paramArray = new object[2];
			paramArray.SetValue(targetNode,0);
			paramArray[1] = swapChildren;
			Invoker.Method(this, "SwapNode", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.DiagramNode CloneNode(bool copyChildren, LateBindingApi.Office.DiagramNode targetNode, LateBindingApi.Office.Enums.MsoRelativeNodePosition pos)
		{
			object[] paramArray = new object[3];
			paramArray[0] = copyChildren;
			paramArray.SetValue(targetNode,1);
			paramArray[2] = pos;
			object returnValue = Invoker.MethodReturn(this, "CloneNode", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.DiagramNode newClass = new LateBindingApi.Office.DiagramNode(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void TransferChildren(LateBindingApi.Office.DiagramNode receivingNode)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(receivingNode,0);
			Invoker.Method(this, "TransferChildren", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.DiagramNode NextNode()
		{
			object returnValue = Invoker.MethodReturn(this, "NextNode", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.DiagramNode newClass = new LateBindingApi.Office.DiagramNode(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.DiagramNode PrevNode()
		{
			object returnValue = Invoker.MethodReturn(this, "PrevNode", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.DiagramNode newClass = new LateBindingApi.Office.DiagramNode(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
