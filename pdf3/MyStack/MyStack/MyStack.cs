using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyStack
{
	class MyStack
	{
		private List<char> elements = new List<char>();
		private int maxLength;
		public MyStack(int Length)
		{
			maxLength = Length;
		}
		public void Push(char Element)
		{
			if (elements.Count != maxLength)
			{
				elements.Add(Element);
			}
			else
			{
				throw new Exception();
			}
		}
		public char Pop()
		{
			if (elements.Count != 0)
			{
				char res = elements[elements.Count() - 1];
				elements.RemoveAt(elements.Count() - 1);
				return res;
			}
			else
			{
				throw new Exception();
			}
		}
		public char Top()
		{
			if (elements.Count != 0)
			{
				return elements[elements.Count() - 1];
			}
			else
			{
				throw new Exception();
			}
		}
		public bool IsEmpty()
		{
			if (elements.Count == 0)
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		public void Clear()
		{
			elements.Clear();
		}
	}
}
