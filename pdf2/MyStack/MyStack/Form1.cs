using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyStack
{
	public partial class Form1 : Form
	{
		MyStack myStack;
		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			try
			{
				myStack = new MyStack(Convert.ToInt32(textBox1.Text));
			}
			catch(FormatException)
			{
				MessageBox.Show("Set the length correctly!");
			}
			textBox1.Clear();
		}

		private void button2_Click(object sender, EventArgs e)
		{
			try 
			{
				myStack.Push(Convert.ToChar(textBox2.Text));
			}
			catch(FormatException)
			{
				MessageBox.Show("Set new element correctly");
			}
			catch
			{
				MessageBox.Show("MyStack overflow");
			}
			textBox2.Clear();
		}

		private void button3_Click(object sender, EventArgs e)
		{
			try
			{
				MessageBox.Show(myStack.Pop().ToString()+" will be removed");
			}
			catch
			{
				MessageBox.Show("Stack is empty");
			}
		}

		private void button4_Click(object sender, EventArgs e)
		{
			try
			{
				MessageBox.Show(myStack.Top().ToString());
			}
			catch
			{
				MessageBox.Show("Stack is empty");
			}
		}

		private void button5_Click(object sender, EventArgs e)
		{
			if (myStack.IsEmpty())
			{
				MessageBox.Show("MyStack is empty");
			}
			else
			{
				MessageBox.Show("MyStack isn't empty");
			}
		}

		private void button6_Click(object sender, EventArgs e)
		{
			myStack.Clear();
		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}
		private string Check()
		{
			MyStack s = new MyStack(textBox3.Text.Length / 2);
			char compare;
			foreach (char symbol in textBox3.Text)
			{
				if (symbol.Equals(Convert.ToChar("{")) || symbol.Equals(Convert.ToChar("(")) || symbol.Equals(Convert.ToChar("["))) s.Push(symbol);
				if (symbol.Equals(Convert.ToChar("}")) || symbol.Equals(Convert.ToChar(")")) || symbol.Equals(Convert.ToChar("]")))
				{
					if (s.IsEmpty()) return ("Incorrect!");
					else compare = s.Top();
					if (symbol.Equals(Convert.ToChar("}")) && compare.Equals(Convert.ToChar("{")) || symbol.Equals(Convert.ToChar(")")) && compare.Equals(Convert.ToChar("(")) || symbol.Equals(Convert.ToChar("]")) && compare.Equals(Convert.ToChar("[")))
					{
						s.Pop();
					}
					else
					{
						return("Incorrect!");
					}
				}
			}
			if (!s.IsEmpty()) return("Incorrect!");
			else return("Correct!");
		}
		private void button7_Click(object sender, EventArgs e)
		{
			MessageBox.Show(Check());
		}
	}
}
