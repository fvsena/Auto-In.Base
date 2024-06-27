using HtmlAgilityPack;
using System;
using System.Collections.Generic;

namespace TPAuto_In.Base.Util
{
    internal class HtmlHelper
    {
        public HtmlDocument document = new HtmlDocument();

        public HtmlHelper(string html)
        {
            document.LoadHtml(html);
        }

        public IEnumerable<HtmlNode> GetElementsByTagName(string tagName)
        {
            if (document == null) throw new Exception("Não é possível interagir com documento nulo");
            return document.DocumentNode.Descendants(tagName);
        }

        public IEnumerable<HtmlNode> GetElementsByTagName(HtmlNode parent, string tagName)
        {
            if (document == null) throw new Exception("Não é possível interagir com documento nulo");
            return parent.Descendants(tagName);
        }

        public HtmlNode GetElementByProperty(IEnumerable<HtmlNode> nodes, string propertyName, string propertyValue)
        {
            HtmlNode n = null;
            foreach (HtmlNode node in nodes)
            {
                HtmlAttributeCollection attributes = node.Attributes;
                if (attributes.Count > 0)
                {
                    foreach (HtmlAttribute att in attributes)
                    {
                        if (att.Name != null && att.Name.Equals(propertyName) && att.Value.Equals(propertyValue))
                        {
                            n = node;
                            break;
                        }
                    }
                    if (n != null)
                    {
                        break;
                    }
                }
            }
            return n;
        }

        public HtmlNode GetElementContainsProperty(IEnumerable<HtmlNode> nodes, string propertyName, string propertyValue)
        {
            HtmlNode n = null;
            foreach (HtmlNode node in nodes)
            {
                HtmlAttributeCollection attributes = node.Attributes;
                if (attributes.Count > 0)
                {
                    foreach (HtmlAttribute att in attributes)
                    {
                        if (att.Name != null && att.Name.Equals(propertyName) && att.Value.Contains(propertyValue))
                        {
                            n = node;
                            break;
                        }
                    }
                    if (n != null)
                    {
                        break;
                    }
                }
            }
            return n;
        }

        public List<HtmlNode> GetElementsContainsProperty(IEnumerable<HtmlNode> nodes, string propertyName, string propertyValue)
        {
            HtmlNode n = null;
            List<HtmlNode> nodesFounded = new List<HtmlNode>();
            foreach (HtmlNode node in nodes)
            {
                HtmlAttributeCollection attributes = node.Attributes;
                if (attributes.Count > 0)
                {
                    foreach (HtmlAttribute att in attributes)
                    {
                        if (att.Name != null && att.Name.Equals(propertyName) && att.Value.Contains(propertyValue))
                        {
                            n = node;
                            nodesFounded.Add(n);
                        }
                    }
                }
            }
            return nodesFounded;
        }

        public List<HtmlNode> GetElementsByProperty(IEnumerable<HtmlNode> nodes, string propertyName, string propertyValue)
        {
            HtmlNode n = null;
            List<HtmlNode> nodesFounded = new List<HtmlNode>();
            foreach (HtmlNode node in nodes)
            {
                HtmlAttributeCollection attributes = node.Attributes;
                if (attributes.Count > 0)
                {
                    foreach (HtmlAttribute att in attributes)
                    {
                        if (att.Name != null && att.Name.Equals(propertyName))
                        {
                            var x = "";
                        }
                        if (att.Name != null && att.Name.Equals(propertyName) && att.Value.Equals(propertyValue))
                        {
                            n = node;
                            nodesFounded.Add(n);
                        }
                    }
                }
            }
            return nodesFounded;
        }

        public List<HtmlNode> GetElementsByTrimProperty(IEnumerable<HtmlNode> nodes, string propertyName, string propertyValue)
        {
            HtmlNode n = null;
            List<HtmlNode> nodesFounded = new List<HtmlNode>();
            foreach (HtmlNode node in nodes)
            {
                HtmlAttributeCollection attributes = node.Attributes;
                if (attributes.Count > 0)
                {
                    foreach (HtmlAttribute att in attributes)
                    {
                        if (att.Name != null && att.Name.Trim().Equals(propertyName.Trim()) && att.Value.Trim().Equals(propertyValue.Trim()))
                        {
                            n = node;
                            nodesFounded.Add(n);
                        }
                    }
                }
            }
            return nodesFounded;
        }

        public HtmlNode GetElementByInnerText(IEnumerable<HtmlNode> nodes, string innerText)
        {
            HtmlNode n = null;
            foreach (HtmlNode node in nodes)
            {
                if (node.InnerText.Equals(innerText))
                {
                    n = node;
                }
            }
            return n;
        }

        public HtmlNode GetElementContainsInnerText(IEnumerable<HtmlNode> nodes, string innerText)
        {
            HtmlNode n = null;
            foreach (HtmlNode node in nodes)
            {
                if (node.InnerText.Contains(innerText))
                {
                    n = node;
                }
            }
            return n;
        }
    }
}
