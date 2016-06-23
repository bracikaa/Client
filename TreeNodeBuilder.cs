using System;
using System.Windows.Forms;
using OpenPop.Mime;
using OpenPop.Mime.Traverse;
using Message = OpenPop.Mime.Message;

namespace cs499
{
    //creating a tree node which looks like a message
	internal class TreeNodeBuilder : IAnswerMessageTraverser<TreeNode>
	{
		public TreeNode VisitMessage(Message message)
		{
			if(message == null)
				throw new ArgumentNullException("message");
			TreeNode child = VisitMessagePart(message.MessagePart);
			TreeNode topNode = new TreeNode(message.Headers.Subject, new[] { child });

			return topNode;
		}

		public TreeNode VisitMessagePart(MessagePart messagePart)
		{
			if(messagePart == null)
				throw new ArgumentNullException("messagePart");
			TreeNode[] children = new TreeNode[0];

			if(messagePart.IsMultiPart)
			{
				children = new TreeNode[messagePart.MessageParts.Count];

				for(int i = 0; i<messagePart.MessageParts.Count; i++)
				{
					children[i] = VisitMessagePart(messagePart.MessageParts[i]);
				}
			}
			TreeNode currentNode = new TreeNode(messagePart.ContentType.MediaType, children);
			currentNode.Tag = messagePart;

			return currentNode;
		}
	}
}