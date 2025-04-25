using System;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ShapeMaster.Services
{
    /// <summary>
    /// Manages all services and their dependencies for the ShapeMaster add-in
    /// </summary>
    public class ServiceManager
    {
        // Private fields for all services
        private readonly PowerPoint.Application _application;
        private NotificationService _notificationService;
        private ShapePositioningService _shapePositioningService;
        private TextFormattingService _textFormattingService;
        private ShapeResizingService _shapeResizingService;
        private EventHandlingService _eventHandlingService;
        private RibbonUIService _ribbonUIService;
        private ErrorHandlingService _errorHandlingService;
        private ComObjectManager _comObjectManager;
        private NoteService _noteService;

        // Singleton instance
        private static ServiceManager _instance;

        // Lock object for thread safety
        private static readonly object _lock = new object();

        /// <summary>
        /// Gets the singleton instance of the ServiceManager
        /// </summary>
        /// <param name="application">PowerPoint application instance (required only on first initialization)</param>
        /// <returns>The singleton ServiceManager instance</returns>
        public static ServiceManager Instance(PowerPoint.Application application = null)
        {
            if (_instance == null)
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        if (application == null)
                        {
                            throw new ArgumentNullException(nameof(application),
                                "PowerPoint application is required for initial ServiceManager creation");
                        }
                        _instance = new ServiceManager(application);
                    }
                }
            }
            return _instance;
        }

        /// <summary>
        /// Private constructor to enforce singleton pattern
        /// </summary>
        /// <param name="application">PowerPoint application instance</param>
        private ServiceManager(PowerPoint.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            InitializeServices();
        }

        /// <summary>
        /// Initializes all services in the correct dependency order
        /// </summary>
        private void InitializeServices()
        {
            try
            {
                // Initialize services in the correct order with dependencies
                _notificationService = new NotificationService();

                // Initialize error handling service early
                _errorHandlingService = new ErrorHandlingService(_notificationService);

                // Initialize COM object manager (with verbose flag set to false by default)
                _comObjectManager = new ComObjectManager(_notificationService.ShowNotification, false);

                // Initialize other services using the notification service and COM object manager
                _shapePositioningService = new ShapePositioningService(_application, _notificationService.ShowNotification, _comObjectManager);
                _textFormattingService = new TextFormattingService(_application, _notificationService.ShowNotification, _comObjectManager);
                _shapeResizingService = new ShapeResizingService(_application, _notificationService.ShowNotification, _errorHandlingService, _comObjectManager);
                _ribbonUIService = new RibbonUIService(_application, _notificationService.ShowNotification, _textFormattingService, _comObjectManager);

                // Initialize the event handling service last, as it depends on other services
                _eventHandlingService = new EventHandlingService(
                    _application,
                    _notificationService.ShowNotification,
                    _textFormattingService,
                    _ribbonUIService.RefreshRibbonUI,
                    _comObjectManager);

                // Subscribe to PowerPoint events
                _eventHandlingService.SubscribeToEvents();

                // Initialize NoteService after other dependencies
                _noteService = new NoteService(_application, _notificationService.ShowNotification);
            }
            catch (Exception ex)
            {
                // Use System.Diagnostics.Debug for logging during initialization
                // as notification service might not be ready
                System.Diagnostics.Debug.WriteLine($"Error initializing services: {ex.Message}");
                if (ex.InnerException != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
                if (!string.IsNullOrEmpty(ex.StackTrace))
                {
                    System.Diagnostics.Debug.WriteLine("Stack trace:");
                    System.Diagnostics.Debug.WriteLine(ex.StackTrace);
                }
                throw new ApplicationException("Failed to initialize services", ex);
            }
        }

        /// <summary>
        /// Properly shuts down all services, unsubscribing from events
        /// </summary>
        public void Shutdown()
        {
            try
            {
                // Unsubscribe from events first
                if (_eventHandlingService != null)
                {
                    _eventHandlingService.UnsubscribeFromEvents();
                }

                // Clean up any cached COM objects in text formatting service
                if (_textFormattingService != null)
                {
                    _textFormattingService.Cleanup();
                }

                // Add cleanup calls for other services if they maintain cached COM objects
                // For example, if ShapeResizingService had a Cleanup method:
                // if (_shapeResizingService != null)
                // {
                //     _shapeResizingService.Cleanup();
                // }

                // Display COM object statistics if the manager was initialized
                if (_comObjectManager != null)
                {
                    System.Diagnostics.Debug.WriteLine("COM Object Release Statistics:");
                    System.Diagnostics.Debug.WriteLine(_comObjectManager.GetReleaseStatistics());
                }
            }
            catch (Exception ex)
            {
                // Log but don't rethrow during shutdown
                System.Diagnostics.Debug.WriteLine($"Error during service shutdown: {ex.Message}");
                if (ex.InnerException != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
            }
        }

        /// <summary>
        /// Sets the ribbon UI for the RibbonUIService
        /// </summary>
        /// <param name="ribbonUI">The IRibbonUI instance</param>
        public void SetRibbonUI(Office.IRibbonUI ribbonUI)
        {
            if (_ribbonUIService != null)
            {
                _ribbonUIService.SetRibbonUI(ribbonUI);
            }
        }

        #region Service Properties

        /// <summary>
        /// Gets the notification service instance
        /// </summary>
        public NotificationService NotificationService => _notificationService;

        /// <summary>
        /// Gets the shape positioning service instance
        /// </summary>
        public ShapePositioningService ShapePositioningService => _shapePositioningService;

        /// <summary>
        /// Gets the text formatting service instance
        /// </summary>
        public TextFormattingService TextFormattingService => _textFormattingService;

        /// <summary>
        /// Gets the shape resizing service instance
        /// </summary>
        public ShapeResizingService ShapeResizingService => _shapeResizingService;

        /// <summary>
        /// Gets the event handling service instance
        /// </summary>
        public EventHandlingService EventHandlingService => _eventHandlingService;

        /// <summary>
        /// Gets the ribbon UI service instance
        /// </summary>
        public RibbonUIService RibbonUIService => _ribbonUIService;

        /// <summary>
        /// Gets the error handling service instance
        /// </summary>
        public ErrorHandlingService ErrorHandlingService => _errorHandlingService;

        /// <summary>
        /// Gets the COM object manager instance
        /// </summary>
        public ComObjectManager ComObjectManager => _comObjectManager;

        /// <summary>
        /// Gets the note service instance
        /// </summary>
        public NoteService NoteService => _noteService;

        #endregion
    }
}