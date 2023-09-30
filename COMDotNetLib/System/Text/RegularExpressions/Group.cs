// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.group?view=netframework-4.8.1

using GRegularExpressions = global::System.Text.RegularExpressions;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DotNetLib.System.Text.RegularExpressions
{
    /// <summary>
    /// Represents the results from a single capturing group.
    /// </summary>
    /// 
    [ComVisible(true)]
    [Guid("88266F44-5087-4D31-A260-57B112072FE9")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.Group")]
    [Description("")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGroup))]
    public class Group : Capture, IGroup, ICapture, IWrappedObject
    {
        internal GRegularExpressions.Group _group;
        internal CaptureCollection _capcoll;

        // Constructors
        internal Group(GRegularExpressions.Group group) : base(group)
        {
            WrappedGroup = group;
        }

        //Properties
        internal GRegularExpressions.Group WrappedGroup
        {
            get { return _group; }
            set 
            { 
                _group = value;
                //_capcoll = new CaptureCollection(value.Captures);
            }  
        }

        //
        // Summary:
        //     Gets a collection of all the captures matched by the capturing group, in innermost-leftmost-first
        //     order (or innermost-rightmost-first order if the regular expression is modified
        //     with the System.Text.RegularExpressions.RegexOptions.RightToLeft option). The
        //     collection may have zero or more items.
        //
        // Returns:
        //     The collection of substrings matched by the group.
        public CaptureCollection Captures
        {
            get
            {
                if (_capcoll == null)
                {
                    _capcoll = new CaptureCollection(this);
                }

                return _capcoll;
            }
        }

        new public object WrappedObject => _group;

        //public CaptureCollection Captures => _capcoll;

        //
        // Summary:
        //     Returns the name of the capturing group represented by the current instance.
        //
        // Returns:
        //     The name of the capturing group represented by the current instance.
        public string Name => _group.Name;


        //
        // Summary:
        //     Gets a value indicating whether the match is successful.
        //
        // Returns:
        //     true if the match is successful; otherwise, false.
        public bool Success => _group.Success;


        //public int Index => _group.Index;

        //public int Length => _group.Length;

        //public string Value => _group.Value;

        public static Group Synchronized(Group inner)
        {
            return new Group(GRegularExpressions.Group.Synchronized(inner.WrappedGroup));
        }
    }
}
