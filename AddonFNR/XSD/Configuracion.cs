﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Este código fue generado por una herramienta.
//     Versión de runtime:4.0.30319.42000
//
//     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
//     se vuelve a generar el código.
// </auto-generated>
//------------------------------------------------------------------------------

using System.Xml.Serialization;

// 
// Este código fuente fue generado automáticamente por xsd, Versión=4.0.30319.33440.
// 


/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
[System.Xml.Serialization.XmlRootAttribute(Namespace="", IsNullable=false)]
public partial class CONFIGURACION {
    
    private CONFIGURACIONPBD pBDField;
    
    private CONFIGURACIONSOCIOSNEGOCIOS sOCIOSNEGOCIOSField;
    
    private CONFIGURACIONBONO bONOField;
    
    private CONFIGURACIONINVERSIONINICIAL iNVERSIONINICIALField;
    
    private CONFIGURACIONSERIECONTRATOS sERIECONTRATOSField;
    
    private CONFIGURACIONCUENTASTRASPASOS cUENTASTRASPASOSField;
    
    /// <remarks/>
    public CONFIGURACIONPBD PBD {
        get {
            return this.pBDField;
        }
        set {
            this.pBDField = value;
        }
    }
    
    /// <remarks/>
    public CONFIGURACIONSOCIOSNEGOCIOS SOCIOSNEGOCIOS {
        get {
            return this.sOCIOSNEGOCIOSField;
        }
        set {
            this.sOCIOSNEGOCIOSField = value;
        }
    }
    
    /// <remarks/>
    public CONFIGURACIONBONO BONO {
        get {
            return this.bONOField;
        }
        set {
            this.bONOField = value;
        }
    }
    
    /// <remarks/>
    public CONFIGURACIONINVERSIONINICIAL INVERSIONINICIAL {
        get {
            return this.iNVERSIONINICIALField;
        }
        set {
            this.iNVERSIONINICIALField = value;
        }
    }
    
    /// <remarks/>
    public CONFIGURACIONSERIECONTRATOS SERIECONTRATOS {
        get {
            return this.sERIECONTRATOSField;
        }
        set {
            this.sERIECONTRATOSField = value;
        }
    }
    
    /// <remarks/>
    public CONFIGURACIONCUENTASTRASPASOS CUENTASTRASPASOS {
        get {
            return this.cUENTASTRASPASOSField;
        }
        set {
            this.cUENTASTRASPASOSField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
public partial class CONFIGURACIONPBD {
    
    private string pBDField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string PBD {
        get {
            return this.pBDField;
        }
        set {
            this.pBDField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
public partial class CONFIGURACIONSOCIOSNEGOCIOS {
    
    private string camposSocioNegociosField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string CamposSocioNegocios {
        get {
            return this.camposSocioNegociosField;
        }
        set {
            this.camposSocioNegociosField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
public partial class CONFIGURACIONBONO {
    
    private string cuentaDebitoBonoField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string CuentaDebitoBono {
        get {
            return this.cuentaDebitoBonoField;
        }
        set {
            this.cuentaDebitoBonoField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
public partial class CONFIGURACIONINVERSIONINICIAL {
    
    private string cuentaInversionInicialField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string CuentaInversionInicial {
        get {
            return this.cuentaInversionInicialField;
        }
        set {
            this.cuentaInversionInicialField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
public partial class CONFIGURACIONSERIECONTRATOS {
    
    private string serieContratosAutomaticaField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string SerieContratosAutomatica {
        get {
            return this.serieContratosAutomaticaField;
        }
        set {
            this.serieContratosAutomaticaField = value;
        }
    }
}

/// <remarks/>
[System.CodeDom.Compiler.GeneratedCodeAttribute("xsd", "4.0.30319.33440")]
[System.SerializableAttribute()]
[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.ComponentModel.DesignerCategoryAttribute("code")]
[System.Xml.Serialization.XmlTypeAttribute(AnonymousType=true)]
public partial class CONFIGURACIONCUENTASTRASPASOS {
    
    private string cuentaApoyoCooperativaField;
    
    private string cuentaCooperativaApoyoField;
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string CuentaApoyoCooperativa {
        get {
            return this.cuentaApoyoCooperativaField;
        }
        set {
            this.cuentaApoyoCooperativaField = value;
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlAttributeAttribute()]
    public string CuentaCooperativaApoyo {
        get {
            return this.cuentaCooperativaApoyoField;
        }
        set {
            this.cuentaCooperativaApoyoField = value;
        }
    }
}
