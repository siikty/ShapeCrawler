﻿using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents an element on a slide.
    /// </summary>
    public abstract class Element
    {
        #region Fields

        protected OpenXmlCompositeElement CompositeElement;

        private bool? _isPlaceholder;
        
        private bool? _hidden;
        private int _id;
        private string _name;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets an element identifier.
        /// </summary>
        public int Id
        {
            get
            {
                InitIdHiddenName();
                return _id;
            }
        }

        /// <summary>
        /// Determines whether the element is hidden.
        /// </summary>
        public bool Hidden
        {
            get
            {
                InitIdHiddenName();
                return (bool)_hidden;
            }
        }

        /// <summary>
        /// Gets an element name.
        /// </summary>
        public string Name
        {
            get
            {
                InitIdHiddenName();
                return _name;
            }
        }

        /// <summary>
        /// Determines whether the element is placeholder.
        /// </summary>
        public bool IsPlaceholder
        {
            get
            {
                if (_isPlaceholder == null)
                {
                    _isPlaceholder = CompositeElement.Descendants<P.PlaceholderShape>().Any();
                }

                return (bool)_isPlaceholder;
            }
        }

        /// <summary>
        /// Gets or sets element type.
        /// </summary>
        public ElementType Type { get; set; } //TODO: remove public modifier for setter

        /// <summary>
        /// Gets or sets the x-coordinate of the upper-left corner of the element in EMUs.
        /// </summary>
        public long X { get; set; }

        /// <summary>
        /// Gets or sets the y-coordinate of the upper-left corner of the element in EMUs.
        /// </summary>
        public long Y { get; set; }

        /// <summary>
        /// Gets or sets width of the element in EMUs.
        /// </summary>
        public long Width { get; set; }

        /// <summary>
        /// Gets or sets height of the element in EMUs.
        /// </summary>
        public long Height { get; set; }

        /// <summary>
        /// Gets or sets tag which can be used for any reason.
        /// </summary>
        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public object Tag { get; set; }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Element"/> class.
        /// </summary>
        protected Element(ElementType et)
        {
            Type = et;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Element"/> class.
        /// </summary>
        protected Element(ElementType et, OpenXmlCompositeElement compositeElement) : this(et)
        {
            Check.NotNull(compositeElement, nameof(compositeElement));
            CompositeElement = compositeElement;
        }

        #endregion Constructors

        #region Private Methods

        private void InitIdHiddenName()
        {
            if (_id == 0) // id == 0: it is mean NonVisualDrawingProperties was not parsed before
            {
                var (id, hidden, name) = CompositeElement.GetNvPrValues();
                _id = id;
                _hidden = hidden;
                _name = name;
            }
        }

        #endregion Private Methods
    }
}