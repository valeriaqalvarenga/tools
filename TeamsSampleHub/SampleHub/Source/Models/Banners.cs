// <copyright file="Banners.cs" company="Microsoft">
// Copyright (c) Microsoft. All Rights Reserved.
// </copyright>

namespace BrandHome.Models
{
    /// <summary>
    /// Banners Model
    /// </summary>
    public class Banners
    {
        /// <summary>
        /// Gets or sets SharePoint Banners component array value
        /// </summary>
        public Value[] Value { get; set; }
    }

    /// <summary>
    /// Banners item model for SharePoint component
    /// </summary>
    public class ValueBanners
    {
        /// <summary>
        /// Gets or sets SharePoint Banners Title
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets SharePoint Banners Description
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets SharePoint Banners Link
        /// </summary>
        public string Image { get; set; }
    }
}