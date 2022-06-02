using System;
using System.Collections.Generic;
using System.Linq;
namespace OpenXmlNbdWrapper
{
    public class NbdHeadline
    {
       
        private int noChildren;
        private NbdHeadline parent;
        private int level;
        public List<NbdHeadline> children;

        public NbdHeadline() {
            children = new List<NbdHeadline>();
            noChildren = 0;
            level = 0;
        }
        public String StyleName { get; set; }
        public String Text { get; set; }

        public int Level { get { return level; } }
        public int NoChildren { get { return noChildren; } }
        public List<NbdHeadline> Children { get { return children; } }
        public String Content { get; set; }

        public void AddChildren(NbdHeadline child) {
            children.Add(child);
            child.Parent = this;
            noChildren += 1;
        }
        public NbdHeadline Parent
        {
            get { return this.parent; }
            set
            {
                this.parent = value;
                this.level = value.Level + 1;
            }
        }
        public String GetPath(NbdHeadline node,String path) {
            String text = "\n" + new string(Enumerable.Range(1, level).Select(i => '*').ToArray()) + Text
                + "\n" + Content;
                ;
            foreach (NbdHeadline child in this.children)
            {
                text += child.GetPath(child, path);
            }
            return text;
        }

        public String GetAllContent(NbdHeadline node, String path)
        {
            String text = "\n" + Text
                        + "\n" + Content;
            ;
            foreach (NbdHeadline child in this.children)
            {
                text += child.GetAllContent(child, path);
            }
            return text;
        }
    }
}
