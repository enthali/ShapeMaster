using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Service for managing and releasing COM objects to prevent memory leaks
    /// </summary>
    public class ComObjectManager
    {
        // Statistics for tracking released objects
        private int _totalObjectsReleased = 0;
        private Dictionary<string, int> _objectTypeReleaseCount = new Dictionary<string, int>();
        private readonly Action<string, bool> _notificationCallback;
        private readonly bool _verbose;

        /// <summary>
        /// Initializes a new instance of the ComObjectManager class
        /// </summary>
        /// <param name="notificationCallback">Optional callback for displaying notifications</param>
        /// <param name="verbose">Whether to log detailed information about released objects</param>
        public ComObjectManager(Action<string, bool> notificationCallback = null, bool verbose = false)
        {
            _notificationCallback = notificationCallback;
            _verbose = verbose;
        }

        /// <summary>
        /// Safely releases a COM object reference
        /// </summary>
        /// <param name="comObject">The COM object to release</param>
        /// <param name="objectName">Optional name for tracking/logging purposes</param>
        /// <returns>True if object was released successfully, false otherwise</returns>
        public bool ReleaseComObject(object comObject, string objectName = null)
        {
            if (comObject == null)
            {
                return false;
            }

            try
            {
                string objectTypeName = comObject.GetType().Name;
                string displayName = objectName ?? objectTypeName;

                int refCount = Marshal.ReleaseComObject(comObject);
                _totalObjectsReleased++;

                // Track counts by type
                if (!_objectTypeReleaseCount.ContainsKey(objectTypeName))
                {
                    _objectTypeReleaseCount[objectTypeName] = 1;
                }
                else
                {
                    _objectTypeReleaseCount[objectTypeName]++;
                }

                // Verbose logging if enabled
                if (_verbose)
                {
                    string message = $"Released {displayName} - Ref count: {refCount}";
                    Debug.WriteLine(message);
                }

                return true;
            }
            catch (Exception ex)
            {
                string errorMessage = $"Error releasing COM object: {ex.Message}";
                Debug.WriteLine(errorMessage);

                if (_notificationCallback != null && _verbose)
                {
                    _notificationCallback(errorMessage, true);
                }

                return false;
            }
        }

        /// <summary>
        /// Safely releases a collection of COM objects
        /// </summary>
        /// <param name="comObjects">Collection of COM objects to release</param>
        /// <param name="collectionName">Optional name for the collection for tracking/logging</param>
        /// <returns>Count of successfully released objects</returns>
        public int ReleaseComObjects(IEnumerable<object> comObjects, string collectionName = null)
        {
            if (comObjects == null)
            {
                return 0;
            }

            int releasedCount = 0;

            foreach (var obj in comObjects)
            {
                if (ReleaseComObject(obj))
                {
                    releasedCount++;
                }
            }

            if (_verbose && releasedCount > 0)
            {
                string message = $"Released {releasedCount} objects from {collectionName ?? "collection"}";
                Debug.WriteLine(message);
            }

            return releasedCount;
        }

        /// <summary>
        /// Safely finalizes a COM object by setting it to null after releasing it
        /// </summary>
        /// <typeparam name="T">Type of the COM object</typeparam>
        /// <param name="comObject">Reference to the COM object to finalize</param>
        /// <param name="objectName">Optional name for tracking/logging purposes</param>
        public void FinalReleaseComObject<T>(ref T comObject, string objectName = null) where T : class
        {
            if (comObject != null)
            {
                ReleaseComObject(comObject, objectName);
                comObject = null;
            }
        }

        /// <summary>
        /// Gets statistics about released COM objects
        /// </summary>
        /// <returns>String containing release statistics</returns>
        public string GetReleaseStatistics()
        {
            string stats = $"Total COM objects released: {_totalObjectsReleased}\n";

            foreach (var type in _objectTypeReleaseCount)
            {
                stats += $"  {type.Key}: {type.Value}\n";
            }

            return stats;
        }

        /// <summary>
        /// Resets the COM object release statistics
        /// </summary>
        public void ResetStatistics()
        {
            _totalObjectsReleased = 0;
            _objectTypeReleaseCount.Clear();
        }
    }
}