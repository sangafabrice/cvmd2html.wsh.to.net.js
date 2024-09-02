namespace ROOT.CIMV2 {
    using System;
    using System.ComponentModel;
    using System.Management;
    using System.Collections;
    using System.Globalization;
    
    
    // Functions ShouldSerialize<PropertyName> are functions used by VS property browser to check if a particular property has to be serialized. These functions are added for all ValueType properties ( properties of type Int32, BOOL etc.. which cannot be set to null). These functions use Is<PropertyName>Null function. These functions are also used in the TypeConverter implementation for the properties to check for NULL value of property so that an empty value can be shown in Property browser in case of Drag and Drop in Visual studio.
    // Functions Is<PropertyName>Null() are used to check if a property is NULL.
    // Functions Reset<PropertyName> are added for Nullable Read/Write properties. These functions are used by VS designer in property browser to set a property to NULL.
    // Every property added to the class for WMI property has attributes set to define its behavior in Visual Studio designer and also to define a TypeConverter to be used.
    // An Early Bound class generated for the WMI class.StdRegProv
    public class StdRegProv : System.ComponentModel.Component {
        
        // Private property to hold the WMI namespace in which the class resides.
        private static string CreatedWmiNamespace = "root\\cimv2";
        
        // Private property to hold the name of WMI class which created this class.
        private static string CreatedClassName = "StdRegProv";
        
        // Private member variable to hold the ManagementScope which is used by the various methods.
        private static System.Management.ManagementScope statMgmtScope = null;
        
        private ManagementSystemProperties PrivateSystemProperties;
        
        // Underlying lateBound WMI object.
        private System.Management.ManagementObject PrivateLateBoundObject;
        
        // Member variable to store the 'automatic commit' behavior for the class.
        private bool AutoCommitProp;
        
        // Private variable to hold the embedded property representing the instance.
        private System.Management.ManagementBaseObject embeddedObj;
        
        // The current WMI object used
        private System.Management.ManagementBaseObject curObj;
        
        // Flag to indicate if the instance is an embedded object.
        private bool isEmbedded;
        
        // Below are different overloads of constructors to initialize an instance of the class with a WMI object.
        public StdRegProv() {
            this.InitializeObject(null, null, null);
        }
        
        public StdRegProv(System.Management.ManagementPath path, System.Management.ObjectGetOptions getOptions) {
            this.InitializeObject(null, path, getOptions);
        }
        
        public StdRegProv(System.Management.ManagementScope mgmtScope, System.Management.ManagementPath path) {
            this.InitializeObject(mgmtScope, path, null);
        }
        
        public StdRegProv(System.Management.ManagementPath path) {
            this.InitializeObject(null, path, null);
        }
        
        public StdRegProv(System.Management.ManagementScope mgmtScope, System.Management.ManagementPath path, System.Management.ObjectGetOptions getOptions) {
            this.InitializeObject(mgmtScope, path, getOptions);
        }
        
        public StdRegProv(System.Management.ManagementObject theObject) {
            Initialize();
            if ((CheckIfProperClass(theObject) == true)) {
                PrivateLateBoundObject = theObject;
                PrivateSystemProperties = new ManagementSystemProperties(PrivateLateBoundObject);
                curObj = PrivateLateBoundObject;
            }
            else {
                throw new System.ArgumentException("Class name does not match.");
            }
        }
        
        public StdRegProv(System.Management.ManagementBaseObject theObject) {
            Initialize();
            if ((CheckIfProperClass(theObject) == true)) {
                embeddedObj = theObject;
                PrivateSystemProperties = new ManagementSystemProperties(theObject);
                curObj = embeddedObj;
                isEmbedded = true;
            }
            else {
                throw new System.ArgumentException("Class name does not match.");
            }
        }
        
        // Property returns the namespace of the WMI class.
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string OriginatingNamespace {
            get {
                return "root\\cimv2";
            }
        }
        
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string ManagementClassName {
            get {
                string strRet = CreatedClassName;
                if ((curObj != null)) {
                    if ((curObj.ClassPath != null)) {
                        strRet = ((string)(curObj["__CLASS"]));
                        if (((strRet == null) 
                                    || (strRet == string.Empty))) {
                            strRet = CreatedClassName;
                        }
                    }
                }
                return strRet;
            }
        }
        
        // Property pointing to an embedded object to get System properties of the WMI object.
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ManagementSystemProperties SystemProperties {
            get {
                return PrivateSystemProperties;
            }
        }
        
        // Property returning the underlying lateBound object.
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public System.Management.ManagementBaseObject LateBoundObject {
            get {
                return curObj;
            }
        }
        
        // ManagementScope of the object.
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public System.Management.ManagementScope Scope {
            get {
                if ((isEmbedded == false)) {
                    return PrivateLateBoundObject.Scope;
                }
                else {
                    return null;
                }
            }
            set {
                if ((isEmbedded == false)) {
                    PrivateLateBoundObject.Scope = value;
                }
            }
        }
        
        // Property to show the commit behavior for the WMI object. If true, WMI object will be automatically saved after each property modification.(ie. Put() is called after modification of a property).
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool AutoCommit {
            get {
                return AutoCommitProp;
            }
            set {
                AutoCommitProp = value;
            }
        }
        
        // The ManagementPath of the underlying WMI object.
        [Browsable(true)]
        public System.Management.ManagementPath Path {
            get {
                if ((isEmbedded == false)) {
                    return PrivateLateBoundObject.Path;
                }
                else {
                    return null;
                }
            }
            set {
                if ((isEmbedded == false)) {
                    if ((CheckIfProperClass(null, value, null) != true)) {
                        throw new System.ArgumentException("Class name does not match.");
                    }
                    PrivateLateBoundObject.Path = value;
                }
            }
        }
        
        // Public static scope property which is used by the various methods.
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public static System.Management.ManagementScope StaticScope {
            get {
                return statMgmtScope;
            }
            set {
                statMgmtScope = value;
            }
        }
        
        private bool CheckIfProperClass(System.Management.ManagementScope mgmtScope, System.Management.ManagementPath path, System.Management.ObjectGetOptions OptionsParam) {
            if (((path != null) 
                        && (string.Compare(path.ClassName, this.ManagementClassName, true, System.Globalization.CultureInfo.InvariantCulture) == 0))) {
                return true;
            }
            else {
                return CheckIfProperClass(new System.Management.ManagementObject(mgmtScope, path, OptionsParam));
            }
        }
        
        private bool CheckIfProperClass(System.Management.ManagementBaseObject theObj) {
            if (((theObj != null) 
                        && (string.Compare(((string)(theObj["__CLASS"])), this.ManagementClassName, true, System.Globalization.CultureInfo.InvariantCulture) == 0))) {
                return true;
            }
            else {
                System.Array parentClasses = ((System.Array)(theObj["__DERIVATION"]));
                if ((parentClasses != null)) {
                    int count = 0;
                    for (count = 0; (count < parentClasses.Length); count = (count + 1)) {
                        if ((string.Compare(((string)(parentClasses.GetValue(count))), this.ManagementClassName, true, System.Globalization.CultureInfo.InvariantCulture) == 0)) {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        
        [Browsable(true)]
        public void CommitObject() {
            if ((isEmbedded == false)) {
                PrivateLateBoundObject.Put();
            }
        }
        
        [Browsable(true)]
        public void CommitObject(System.Management.PutOptions putOptions) {
            if ((isEmbedded == false)) {
                PrivateLateBoundObject.Put(putOptions);
            }
        }
        
        private void Initialize() {
            AutoCommitProp = true;
            isEmbedded = false;
        }
        
        private static string ConstructPath() {
            string strPath = "root\\cimv2:StdRegProv";
            return strPath;
        }
        
        private void InitializeObject(System.Management.ManagementScope mgmtScope, System.Management.ManagementPath path, System.Management.ObjectGetOptions getOptions) {
            Initialize();
            if ((path != null)) {
                if ((CheckIfProperClass(mgmtScope, path, getOptions) != true)) {
                    throw new System.ArgumentException("Class name does not match.");
                }
            }
            PrivateLateBoundObject = new System.Management.ManagementObject(mgmtScope, path, getOptions);
            PrivateSystemProperties = new ManagementSystemProperties(PrivateLateBoundObject);
            curObj = PrivateLateBoundObject;
        }
        
        // Different overloads of GetInstances() help in enumerating instances of the WMI class.
        public static StdRegProvCollection GetInstances() {
            return GetInstances(null, null, null);
        }
        
        public static StdRegProvCollection GetInstances(string condition) {
            return GetInstances(null, condition, null);
        }
        
        public static StdRegProvCollection GetInstances(string[] selectedProperties) {
            return GetInstances(null, null, selectedProperties);
        }
        
        public static StdRegProvCollection GetInstances(string condition, string[] selectedProperties) {
            return GetInstances(null, condition, selectedProperties);
        }
        
        public static StdRegProvCollection GetInstances(System.Management.ManagementScope mgmtScope, System.Management.EnumerationOptions enumOptions) {
            if ((mgmtScope == null)) {
                if ((statMgmtScope == null)) {
                    mgmtScope = new System.Management.ManagementScope();
                    mgmtScope.Path.NamespacePath = "root\\cimv2";
                }
                else {
                    mgmtScope = statMgmtScope;
                }
            }
            System.Management.ManagementPath pathObj = new System.Management.ManagementPath();
            pathObj.ClassName = "StdRegProv";
            pathObj.NamespacePath = "root\\cimv2";
            System.Management.ManagementClass clsObject = new System.Management.ManagementClass(mgmtScope, pathObj, null);
            if ((enumOptions == null)) {
                enumOptions = new System.Management.EnumerationOptions();
                enumOptions.EnsureLocatable = true;
            }
            return new StdRegProvCollection(clsObject.GetInstances(enumOptions));
        }
        
        public static StdRegProvCollection GetInstances(System.Management.ManagementScope mgmtScope, string condition) {
            return GetInstances(mgmtScope, condition, null);
        }
        
        public static StdRegProvCollection GetInstances(System.Management.ManagementScope mgmtScope, string[] selectedProperties) {
            return GetInstances(mgmtScope, null, selectedProperties);
        }
        
        public static StdRegProvCollection GetInstances(System.Management.ManagementScope mgmtScope, string condition, string[] selectedProperties) {
            if ((mgmtScope == null)) {
                if ((statMgmtScope == null)) {
                    mgmtScope = new System.Management.ManagementScope();
                    mgmtScope.Path.NamespacePath = "root\\cimv2";
                }
                else {
                    mgmtScope = statMgmtScope;
                }
            }
            System.Management.ManagementObjectSearcher ObjectSearcher = new System.Management.ManagementObjectSearcher(mgmtScope, new SelectQuery("StdRegProv", condition, selectedProperties));
            System.Management.EnumerationOptions enumOptions = new System.Management.EnumerationOptions();
            enumOptions.EnsureLocatable = true;
            ObjectSearcher.Options = enumOptions;
            return new StdRegProvCollection(ObjectSearcher.Get());
        }
        
        [Browsable(true)]
        public static StdRegProv CreateInstance() {
            System.Management.ManagementScope mgmtScope = null;
            if ((statMgmtScope == null)) {
                mgmtScope = new System.Management.ManagementScope();
                mgmtScope.Path.NamespacePath = CreatedWmiNamespace;
            }
            else {
                mgmtScope = statMgmtScope;
            }
            System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
            System.Management.ManagementClass tmpMgmtClass = new System.Management.ManagementClass(mgmtScope, mgmtPath, null);
            return new StdRegProv(tmpMgmtClass.CreateInstance());
        }
        
        [Browsable(true)]
        public void Delete() {
            PrivateLateBoundObject.Delete();
        }
        
        public static uint CheckAccess(uint hDefKey, string sSubKeyName, uint uRequired, out bool bGranted) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("CheckAccess");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["uRequired"] = ((uint)(uRequired));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("CheckAccess", inParams, null);
                bGranted = System.Convert.ToBoolean(outParams.Properties["bGranted"].Value);
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                bGranted = System.Convert.ToBoolean(0);
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint CreateKey(uint hDefKey, string sSubKeyName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("CreateKey");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("CreateKey", inParams, null);
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint DeleteKey(uint hDefKey, string sSubKeyName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("DeleteKey");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("DeleteKey", inParams, null);
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint DeleteValue(uint hDefKey, string sSubKeyName, string sValueName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("DeleteValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("DeleteValue", inParams, null);
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static string[] EnumKey(uint hDefKey, string sSubKeyName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("EnumKey");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("EnumKey", inParams, null);
                return ((string[])(outParams.Properties["sNames"].Value));
            }
            else {
                return null;
            }
        }
        
        public static uint EnumValues(uint hDefKey, string sSubKeyName, out string[] sNames, out int[] Types) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("EnumValues");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("EnumValues", inParams, null);
                sNames = ((string[])(outParams.Properties["sNames"].Value));
                Types = ((int[])(outParams.Properties["Types"].Value));
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                sNames = null;
                Types = null;
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint GetBinaryValue(uint hDefKey, string sSubKeyName, string sValueName, out byte[] uValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("GetBinaryValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("GetBinaryValue", inParams, null);
                uValue = ((byte[])(outParams.Properties["uValue"].Value));
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                uValue = null;
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint GetDWORDValue(uint hDefKey, string sSubKeyName, string sValueName, out uint uValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("GetDWORDValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("GetDWORDValue", inParams, null);
                uValue = System.Convert.ToUInt32(outParams.Properties["uValue"].Value);
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                uValue = System.Convert.ToUInt32(0);
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint GetExpandedStringValue(uint hDefKey, string sSubKeyName, string sValueName, out string sValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("GetExpandedStringValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("GetExpandedStringValue", inParams, null);
                sValue = System.Convert.ToString(outParams.Properties["sValue"].Value);
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                sValue = null;
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint GetMultiStringValue(uint hDefKey, string sSubKeyName, string sValueName, out string[] sValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("GetMultiStringValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("GetMultiStringValue", inParams, null);
                sValue = ((string[])(outParams.Properties["sValue"].Value));
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                sValue = null;
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint GetQWORDValue(uint hDefKey, string sSubKeyName, string sValueName, out ulong uValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                inParams = classObj.GetMethodParameters("GetQWORDValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("GetQWORDValue", inParams, null);
                uValue = System.Convert.ToUInt64(outParams.Properties["uValue"].Value);
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                uValue = System.Convert.ToUInt64(0);
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint GetSecurityDescriptor(uint hDefKey, string sSubKeyName, out System.Management.ManagementBaseObject Descriptor) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("GetSecurityDescriptor");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("GetSecurityDescriptor", inParams, null);
                Descriptor = ((System.Management.ManagementBaseObject)(outParams.Properties["Descriptor"].Value));
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                Descriptor = null;
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static string GetStringValue(uint hDefKey, string sSubKeyName, string sValueName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("GetStringValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("GetStringValue", inParams, null);
                return System.Convert.ToString(outParams.Properties["sValue"].Value);
            }
            else {
                return null;
            }
        }
        
        public static uint SetBinaryValue(uint hDefKey, string sSubKeyName, string sValueName, byte[] uValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("SetBinaryValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                inParams["uValue"] = ((byte[])(uValue));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("SetBinaryValue", inParams, null);
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint SetDWORDValue(uint hDefKey, string sSubKeyName, string sValueName, uint uValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("SetDWORDValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                inParams["uValue"] = ((uint)(uValue));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("SetDWORDValue", inParams, null);
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint SetExpandedStringValue(uint hDefKey, string sSubKeyName, string sValue, string sValueName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("SetExpandedStringValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValue"] = ((string)(sValue));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("SetExpandedStringValue", inParams, null);
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint SetMultiStringValue(uint hDefKey, string sSubKeyName, string[] sValue, string sValueName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("SetMultiStringValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValue"] = ((string[])(sValue));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("SetMultiStringValue", inParams, null);
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint SetQWORDValue(uint hDefKey, string sSubKeyName, string sValueName, ulong uValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("SetQWORDValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValueName"] = ((string)(sValueName));
                inParams["uValue"] = ((ulong)(uValue));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("SetQWORDValue", inParams, null);
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint SetSecurityDescriptor(System.Management.ManagementBaseObject Descriptor, uint hDefKey, string sSubKeyName) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("SetSecurityDescriptor");
                inParams["Descriptor"] = ((System.Management.ManagementBaseObject )(Descriptor));
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("SetSecurityDescriptor", inParams, null);
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        public static uint SetStringValue(uint hDefKey, string sSubKeyName, string sValueName, string sValue) {
            bool IsMethodStatic = true;
            if ((IsMethodStatic == true)) {
                System.Management.ManagementBaseObject inParams = null;
                System.Management.ManagementPath mgmtPath = new System.Management.ManagementPath(CreatedClassName);
                System.Management.ManagementClass classObj = new System.Management.ManagementClass(statMgmtScope, mgmtPath, null);
                bool EnablePrivileges = classObj.Scope.Options.EnablePrivileges;
                classObj.Scope.Options.EnablePrivileges = true;
                inParams = classObj.GetMethodParameters("SetStringValue");
                inParams["hDefKey"] = ((uint)(hDefKey));
                inParams["sSubKeyName"] = ((string)(sSubKeyName));
                inParams["sValue"] = ((string)(sValue));
                inParams["sValueName"] = ((string)(sValueName));
                System.Management.ManagementBaseObject outParams = classObj.InvokeMethod("SetStringValue", inParams, null);
                classObj.Scope.Options.EnablePrivileges = EnablePrivileges;
                return System.Convert.ToUInt32(outParams.Properties["ReturnValue"].Value);
            }
            else {
                return System.Convert.ToUInt32(0);
            }
        }
        
        // Enumerator implementation for enumerating instances of the class.
        public class StdRegProvCollection : object, ICollection {
            
            private ManagementObjectCollection privColObj;
            
            public StdRegProvCollection(ManagementObjectCollection objCollection) {
                privColObj = objCollection;
            }
            
            public virtual int Count {
                get {
                    return privColObj.Count;
                }
            }
            
            public virtual bool IsSynchronized {
                get {
                    return privColObj.IsSynchronized;
                }
            }
            
            public virtual object SyncRoot {
                get {
                    return this;
                }
            }
            
            public virtual void CopyTo(System.Array array, int index) {
                privColObj.CopyTo(array, index);
                int nCtr;
                for (nCtr = 0; (nCtr < array.Length); nCtr = (nCtr + 1)) {
                    array.SetValue(new StdRegProv(((System.Management.ManagementObject)(array.GetValue(nCtr)))), nCtr);
                }
            }
            
            public virtual System.Collections.IEnumerator GetEnumerator() {
                return new StdRegProvEnumerator(privColObj.GetEnumerator());
            }
            
            public class StdRegProvEnumerator : object, System.Collections.IEnumerator {
                
                private ManagementObjectCollection.ManagementObjectEnumerator privObjEnum;
                
                public StdRegProvEnumerator(ManagementObjectCollection.ManagementObjectEnumerator objEnum) {
                    privObjEnum = objEnum;
                }
                
                public virtual object Current {
                    get {
                        return new StdRegProv(((System.Management.ManagementObject)(privObjEnum.Current)));
                    }
                }
                
                public virtual bool MoveNext() {
                    return privObjEnum.MoveNext();
                }
                
                public virtual void Reset() {
                    privObjEnum.Reset();
                }
            }
        }
        
        // TypeConverter to handle null values for ValueType properties
        public class WMIValueTypeConverter : TypeConverter {
            
            private TypeConverter baseConverter;
            
            private System.Type baseType;
            
            public WMIValueTypeConverter(System.Type inBaseType) {
                baseConverter = TypeDescriptor.GetConverter(inBaseType);
                baseType = inBaseType;
            }
            
            public override bool CanConvertFrom(System.ComponentModel.ITypeDescriptorContext context, System.Type srcType) {
                return baseConverter.CanConvertFrom(context, srcType);
            }
            
            public override bool CanConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Type destinationType) {
                return baseConverter.CanConvertTo(context, destinationType);
            }
            
            public override object ConvertFrom(System.ComponentModel.ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value) {
                return baseConverter.ConvertFrom(context, culture, value);
            }
            
            public override object CreateInstance(System.ComponentModel.ITypeDescriptorContext context, System.Collections.IDictionary dictionary) {
                return baseConverter.CreateInstance(context, dictionary);
            }
            
            public override bool GetCreateInstanceSupported(System.ComponentModel.ITypeDescriptorContext context) {
                return baseConverter.GetCreateInstanceSupported(context);
            }
            
            public override PropertyDescriptorCollection GetProperties(System.ComponentModel.ITypeDescriptorContext context, object value, System.Attribute[] attributeVar) {
                return baseConverter.GetProperties(context, value, attributeVar);
            }
            
            public override bool GetPropertiesSupported(System.ComponentModel.ITypeDescriptorContext context) {
                return baseConverter.GetPropertiesSupported(context);
            }
            
            public override System.ComponentModel.TypeConverter.StandardValuesCollection GetStandardValues(System.ComponentModel.ITypeDescriptorContext context) {
                return baseConverter.GetStandardValues(context);
            }
            
            public override bool GetStandardValuesExclusive(System.ComponentModel.ITypeDescriptorContext context) {
                return baseConverter.GetStandardValuesExclusive(context);
            }
            
            public override bool GetStandardValuesSupported(System.ComponentModel.ITypeDescriptorContext context) {
                return baseConverter.GetStandardValuesSupported(context);
            }
            
            public override object ConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, System.Type destinationType) {
                if ((baseType.BaseType == typeof(System.Enum))) {
                    if ((value.GetType() == destinationType)) {
                        return value;
                    }
                    if ((((value == null) 
                                && (context != null)) 
                                && (context.PropertyDescriptor.ShouldSerializeValue(context.Instance) == false))) {
                        return  "NULL_ENUM_VALUE" ;
                    }
                    return baseConverter.ConvertTo(context, culture, value, destinationType);
                }
                if (((baseType == typeof(bool)) 
                            && (baseType.BaseType == typeof(System.ValueType)))) {
                    if ((((value == null) 
                                && (context != null)) 
                                && (context.PropertyDescriptor.ShouldSerializeValue(context.Instance) == false))) {
                        return "";
                    }
                    return baseConverter.ConvertTo(context, culture, value, destinationType);
                }
                if (((context != null) 
                            && (context.PropertyDescriptor.ShouldSerializeValue(context.Instance) == false))) {
                    return "";
                }
                return baseConverter.ConvertTo(context, culture, value, destinationType);
            }
        }
        
        // Embedded class to represent WMI system Properties.
        [TypeConverter(typeof(System.ComponentModel.ExpandableObjectConverter))]
        public class ManagementSystemProperties {
            
            private System.Management.ManagementBaseObject PrivateLateBoundObject;
            
            public ManagementSystemProperties(System.Management.ManagementBaseObject ManagedObject) {
                PrivateLateBoundObject = ManagedObject;
            }
            
            [Browsable(true)]
            public int GENUS {
                get {
                    return ((int)(PrivateLateBoundObject["__GENUS"]));
                }
            }
            
            [Browsable(true)]
            public string CLASS {
                get {
                    return ((string)(PrivateLateBoundObject["__CLASS"]));
                }
            }
            
            [Browsable(true)]
            public string SUPERCLASS {
                get {
                    return ((string)(PrivateLateBoundObject["__SUPERCLASS"]));
                }
            }
            
            [Browsable(true)]
            public string DYNASTY {
                get {
                    return ((string)(PrivateLateBoundObject["__DYNASTY"]));
                }
            }
            
            [Browsable(true)]
            public string RELPATH {
                get {
                    return ((string)(PrivateLateBoundObject["__RELPATH"]));
                }
            }
            
            [Browsable(true)]
            public int PROPERTY_COUNT {
                get {
                    return ((int)(PrivateLateBoundObject["__PROPERTY_COUNT"]));
                }
            }
            
            [Browsable(true)]
            public string[] DERIVATION {
                get {
                    return ((string[])(PrivateLateBoundObject["__DERIVATION"]));
                }
            }
            
            [Browsable(true)]
            public string SERVER {
                get {
                    return ((string)(PrivateLateBoundObject["__SERVER"]));
                }
            }
            
            [Browsable(true)]
            public string NAMESPACE {
                get {
                    return ((string)(PrivateLateBoundObject["__NAMESPACE"]));
                }
            }
            
            [Browsable(true)]
            public string PATH {
                get {
                    return ((string)(PrivateLateBoundObject["__PATH"]));
                }
            }
        }
    }
}
