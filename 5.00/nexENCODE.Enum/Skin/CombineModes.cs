namespace nexENCODE.Enum.Skin {
    /// <summary>
    /// Combine Modes
    /// </summary>
    public enum CombineModes {
        /// <summary>
        /// None
        /// </summary>
        Rgn_None = 0,
        /// <summary>
        /// Creates the intersection of the two combined regions.
        /// </summary>
        Rgn_And = 1,
        /// <summary>
        /// Rgn_Or - Creates a copy of the region identified by hrgnSrc1.
        /// </summary>
        Rgn_Or = 2,
        /// <summary>
        /// Combines the parts of hrgnSrc1 that are not part of hrgnSrc2.
        /// </summary>
        Rgn_XOr = 3,
        /// <summary>
        /// Creates the union of two combined regions.
        /// </summary>
        Rgn_Diff = 4,
        /// <summary>
        /// Creates the union of two combined regions except for any overlapping areas.
        /// </summary>
        Rgn_Copy = 5 
    }
}