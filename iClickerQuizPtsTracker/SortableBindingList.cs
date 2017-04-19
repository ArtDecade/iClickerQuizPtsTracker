using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;

namespace iClickerQuizPtsTracker
{
    /// <summary>
    /// Provides a generic collection that supports both data binding and sorting.
    /// </summary>
    /// <typeparam name="Session"><see cref="iClickerQuizPtsTracker.Session"/>
    /// </typeparam>
    /// <remarks>Adapted from 
    /// <seealso cref="http://martinwilley.com/net/code/forms/sortablebindinglist.html"/>.
    /// <para>That article in turn cites 
    /// <seealso cref="http://msdn.microsoft.com/en-us/library/ms993236.aspx"/>
    /// </para></remarks>
    public class SortableBindingList<Session> : BindingList<Session>
    {
        #region fields
        private bool _isSorted;
        private ListSortDirection _sortDirection = ListSortDirection.Ascending;
        private PropertyDescriptor _sortProperty;
        #endregion

        #region ctors
        /// <summary>
        /// Initializes a new instance of the 
        /// <see cref="SortableBindingList{iClickerQuizPtsTracker.Session}"/> class.
        /// </summary>
        public SortableBindingList()
        {
        }


        //An <see cref="T:System.Collections.Generic.IList`1" /> 
        //of items to be contained in the <see cref="T:System.ComponentModel.BindingList`1" />.
        /// <summary>
        /// Initializes a new instance of the 
        /// <see cref="SortableBindingList{iClickerQuizPtsTracker.Session}"/> class.
        /// </summary>
        /// <param name="list">An <see cref="System.Collections.Generic.IList{Session}"/> 
        /// to be contained in the <see cref="System.ComponentModel.BindingList{Session}"/>.</param>
        public SortableBindingList(IList<Session> list) : base(list)
        {
        }
        #endregion

        #region pptys
        /// <summary>
        /// Gets a value indicating whether the list supports sorting.
        /// </summary>
        protected override bool SupportsSortingCore
        {
            get { return true; }
        }

        /// <summary>
        /// Gets a value indicating whether the list is sorted.
        /// </summary>
        protected override bool IsSortedCore
        {
            get { return _isSorted; }
        }

        /// <summary>
        /// Gets the direction the list is sorted.
        /// </summary>
        protected override ListSortDirection SortDirectionCore
        {
            get { return _sortDirection; }
        }

        /// <summary>
        /// Gets the property descriptor that is used for sorting the list if sorting 
        /// is implemented in a derived class; otherwise returns null.
        /// </summary>
        protected override PropertyDescriptor SortPropertyCore
        {
            get { return _sortProperty; }
        }
        #endregion

        #region methods
        /// <summary>
        /// Removes any sort applied with ApplySortCore if sorting is implemented.
        /// </summary>
        protected override void RemoveSortCore()
        {
            _sortDirection = ListSortDirection.Ascending;
            _sortProperty = null;
            _isSorted = false; 
        }

        /// <summary>
        /// Sorts the items if overridden in a derived class
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="direction"></param>
        protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
        {
            _sortProperty = prop;
            _sortDirection = direction;

            List<Session> list = Items as List<Session>;
            if (list == null) return;

            list.Sort(Compare);

            _isSorted = true;
            //fire an event that the list has been changed.
            OnListChanged(new ListChangedEventArgs(ListChangedType.Reset, -1));
        }

        private int Compare(Session lhs, Session rhs)
        {
            var result = OnComparison(lhs, rhs);
            //invert if descending
            if (_sortDirection == ListSortDirection.Descending)
                result = -result;
            return result;
        }

        private int OnComparison(Session lhs, Session rhs)
        {
            object lhsValue = lhs == null ? null : _sortProperty.GetValue(lhs);
            object rhsValue = rhs == null ? null : _sortProperty.GetValue(rhs);
            return ((IComparable)lhs).CompareTo(rhs);
        }
        #endregion





    }
}
