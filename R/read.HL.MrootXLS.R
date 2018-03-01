#' Reads Hydrolight multi-wavelenght EXCELL spreadsheet (Mroot.xls)
#'
#'@param file is the file name of the multi-wavelenght EXCEL spreadsheet (Mroot.xls)
#'
#'@return  Returns a long list of parameter output by Hydrolight
#'
#'@author Simon Bélanger

read.HL.MrootXLS <- function(file) {

  # Check file extension
  extension <- file_ext(file)
  if (extension != "xlsx" && extension != "xls") {
    print("Not an EXCEL file!")
    return(0)
  }


  if (extension == "xlsx") {
    # Absorption
    #####
    a <- read_xlsx(file, sheet = "a", skip=3)
    x <- colnames(a)
    depth <- as.numeric(x[2:length(x)])
    waves <- as.numeric(a$wavelen)
    a     <- as.matrix(a[,2:length(x)])

    # Only keep the total absorption coefficient
    ixNA  <- which(is.na(waves))
    nwaves<- ixNA[1] - 1
    waves <- waves[1:nwaves]
    a     <- a[1:nwaves,]

    # Scattering
    #####
    b <- read_xlsx(file, sheet = "b", skip=3)
    b <- as.matrix(b[,2:length(x)])
    b <- b[1:nwaves,]

    # Backscattering
    #####
    bb <- read_xlsx(file, sheet = "bb", skip=3)
    bb <- as.matrix(bb[,2:length(x)])
    bb <- bb[1:nwaves,]

    # Backscattering fraction (total)
    #####
    bb.tilde <- read_xlsx(file, sheet = "bb fraction", skip=3)
    bb.tilde <- as.matrix(bb.tilde[,2:length(x)])
    bb.tilde <- bb.tilde[1:nwaves,]

    # Ed 0+
    #####
    Ed.0p <- read_xlsx(file, sheet = "Ed_in_air", skip=3)
    Ed.0p <- Ed.0p[,-1]

    # Lu
    #####
    LuZ <- read_xlsx(file, sheet = "Lu", skip=3)
    x <- colnames(LuZ)
    Lu.0p  <- LuZ$`in air`
    LuZ <- as.matrix(LuZ[,3:length(x)])
    LuZ <- LuZ[1:nwaves,]

    # EdoZ
    #####
    EdoZ <- read_xlsx(file, sheet = "Eod", skip=3)
    x <- colnames(EdoZ)
    Edo.0p  <- EdoZ$`in air`
    EdoZ <- as.matrix(EdoZ[,3:length(x)])
    EdoZ <- EdoZ[1:nwaves,]

    # EuoZ
    #####
    EuoZ <- read_xlsx(file, sheet = "Eou", skip=3)
    x <- colnames(EuoZ)
    Euo.0p  <- EuoZ$`in air`
    EuoZ <- as.matrix(EuoZ[,3:length(x)])
    EuoZ <- EuoZ[1:nwaves,]

    # EoZ
    #####
    EoZ <- read_xlsx(file, sheet = "Eo", skip=3)
    x <- colnames(EoZ)
    Eo.0p  <- EoZ$`in air`
    EoZ <- as.matrix(EoZ[,3:length(x)])
    EoZ <- EoZ[1:nwaves,]

    # QoZ
    #####
    QoZ <- read_xlsx(file, sheet = "Eo_quantum", skip=3)
    x <- colnames(QoZ)
    Qo.0p  <- QoZ$`in air`
    QoZ <- as.matrix(QoZ[,3:length(x)])
    QoZ <- QoZ[1:nwaves,]

    # EdZ
    #####
    EdZ <- read_xlsx(file, sheet = "Ed", skip=3)
    x <- colnames(EdZ)
    Ed.0p  <- EdZ$`in air`
    EdZ <- as.matrix(EdZ[,3:length(x)])
    EdZ <- EdZ[1:nwaves,]

    # EuZ
    #####
    EuZ <- read_xlsx(file, sheet = "Eu", skip=3)
    x <- colnames(EuZ)
    Eu.0p  <- EuZ$`in air`
    EuZ <- as.matrix(EuZ[,3:length(x)])
    EuZ <- EuZ[1:nwaves,]

    # R
    R = EuZ/EdZ
    R.0p = Eu.0p/Ed.0p

    x=read_xlsx(file, sheet = "Rrs", skip=3)
    names(x) <- c("waves", "Rrs"    ,  "Ed"      ,   "Lw"     ,   "Lu")
    Rrs.0p <- x$Rrs
    Lw     <- x$Lw

    # Kd
    #####
    KdZ <- read_xlsx(file, sheet = "Kd", skip=3)
    x <- colnames(KdZ)
    KdZ <- as.matrix(KdZ[,2:length(x)])
    KdZ <- KdZ[1:nwaves,]

    # Ku
    #####
    KuZ <- read_xlsx(file, sheet = "Ku", skip=3)
    KuZ <- as.matrix(KuZ[,2:length(x)])
    KuZ <- KuZ[1:nwaves,]

    # KLu
    #####
    KLuZ <- read_xlsx(file, sheet = "KLu", skip=3)
    KLuZ <- as.matrix(KLuZ[,2:length(x)])
    KLuZ <- KLuZ[1:nwaves,]


    # PAR
    #####
    PAR <- read_xlsx(file, sheet = "PAR", skip=3)
    x <- colnames(PAR)
    PAR <- as.matrix(PAR[,2:length(x)])
    PAR.0p <- PAR[1,]
    PAR.Z <- PAR[-1,]


    # KPAR
    #####
    KPAR <- read_xlsx(file, sheet = "KPAR", skip=3)
    KPAR.Eo <- KPAR$`K_PAR (from Eo)`

  }


  if (extension == "xls")  {
    # Absorption
    #####
    a <- read_xls(file, sheet = "a", skip=3)
    x <- colnames(a)
    depth <- as.numeric(x[2:length(x)])
    waves <- as.numeric(a$wavelen)
    a     <- as.matrix(a[,2:length(x)])

    # Only keep the total absorption coefficient
    ixNA  <- which(is.na(waves))
    nwaves<- ixNA[1] - 1
    waves <- waves[1:nwaves]
    a     <- a[1:nwaves,]

    # Scattering
    #####
    b <- read_xls(file, sheet = "b", skip=3)
    b <- as.matrix(b[,2:length(x)])
    b <- b[1:nwaves,]

    # Backscattering
    #####
    bb <- read_xls(file, sheet = "bb", skip=3)
    bb <- as.matrix(bb[,2:length(x)])
    bb <- bb[1:nwaves,]

    # Backscattering fraction (total)
    #####
    bb.tilde <- read_xls(file, sheet = "bb fraction", skip=3)
    bb.tilde <- as.matrix(bb.tilde[,2:length(x)])
    bb.tilde <- bb.tilde[1:nwaves,]

    # Ed 0+
    #####
    Ed.0p <- read_xls(file, sheet = "Ed_in_air", skip=3)
    Ed.0p <- Ed.0p[,-1]

    # Lu
    #####
    LuZ <- read_xls(file, sheet = "Lu", skip=3)
    x <- colnames(LuZ)
    Lu.0p  <- LuZ$`in air`
    LuZ <- as.matrix(LuZ[,3:length(x)])
    LuZ <- LuZ[1:nwaves,]

    # EdoZ
    #####
    EdoZ <- read_xls(file, sheet = "Eod", skip=3)
    x <- colnames(EdoZ)
    Edo.0p  <- EdoZ$`in air`
    EdoZ <- as.matrix(EdoZ[,3:length(x)])
    EdoZ <- EdoZ[1:nwaves,]

    # EuoZ
    #####
    EuoZ <- read_xls(file, sheet = "Eou", skip=3)
    x <- colnames(EuoZ)
    Euo.0p  <- EuoZ$`in air`
    EuoZ <- as.matrix(EuoZ[,3:length(x)])
    EuoZ <- EuoZ[1:nwaves,]

    # EoZ
    #####
    EoZ <- read_xls(file, sheet = "Eo", skip=3)
    x <- colnames(EoZ)
    Eo.0p  <- EoZ$`in air`
    EoZ <- as.matrix(EoZ[,3:length(x)])
    EoZ <- EoZ[1:nwaves,]

    # QoZ
    #####
    QoZ <- read_xls(file, sheet = "Eo_quantum", skip=3)
    x <- colnames(QoZ)
    Qo.0p  <- QoZ$`in air`
    QoZ <- as.matrix(QoZ[,3:length(x)])
    QoZ <- QoZ[1:nwaves,]

    # EdZ
    #####
    EdZ <- read_xls(file, sheet = "Ed", skip=3)
    x <- colnames(EdZ)
    Ed.0p  <- EdZ$`in air`
    EdZ <- as.matrix(EdZ[,3:length(x)])
    EdZ <- EdZ[1:nwaves,]

    # EuZ
    #####
    EuZ <- read_xls(file, sheet = "Eu", skip=3)
    x <- colnames(EuZ)
    Eu.0p  <- EuZ$`in air`
    EuZ <- as.matrix(EuZ[,3:length(x)])
    EuZ <- EuZ[1:nwaves,]

    # R
    R = EuZ/EdZ
    R.0p = Eu.0p/Ed.0p

    x=read_xls(file, sheet = "Rrs", skip=3)
    names(x) <- c("waves", "Rrs"    ,  "Ed"      ,   "Lw"     ,   "Lu")
    Rrs.0p <- x$Rrs
    Lw     <- x$Lw

    # Kd
    #####
    KdZ <- read_xls(file, sheet = "Kd", skip=3)
    x <- colnames(KdZ)
    KdZ <- as.matrix(KdZ[,2:length(x)])
    KdZ <- KdZ[1:nwaves,]

    # Ku
    #####
    KuZ <- read_xls(file, sheet = "Ku", skip=3)
    KuZ <- as.matrix(KuZ[,2:length(x)])
    KuZ <- KuZ[1:nwaves,]

    # KLu
    #####
    KLuZ <- read_xls(file, sheet = "KLu", skip=3)
    KLuZ <- as.matrix(KLuZ[,2:length(x)])
    KLuZ <- KLuZ[1:nwaves,]


    # PAR
    #####
    PAR <- read_xls(file, sheet = "PAR", skip=3)
    x <- colnames(PAR)
    PAR <- as.matrix(PAR[,2:length(x)])
    PAR.0p <- PAR[1,]
    PAR.Z <- PAR[-1,]


    # KPAR
    #####
    KPAR <- read_xls(file, sheet = "KPAR", skip=3)
    KPAR.Eo <- KPAR$`K_PAR (from Eo)`

  }





  return(list(depth=depth, waves=waves,
              a=a,b=b,bb=bb,
              EoZ=EoZ,
              EdoZ=EdoZ,
              EuoZ=EuoZ,
              EdZ=EdZ,
              EuZ=EuZ,
              KdZ=KdZ,
              KLuZ=KLuZ,
              KuZ=KuZ,
              KPAR.Eo=KPAR.Eo,
              LuZ=LuZ,
              PAR.Z=PAR.Z,
              PAR.0p=PAR.0p,
              QoZ=QoZ,
              Qo.0p=Qo.0p,
              R=R,
              R.0p=R.0p,
              Rrs.0p=Rrs.0p,
              Lw=Lw,
              Ed.0p=Ed.0p,
              Edo.0p=Edo.0p,
              Eo.0p=Eo.0p,
              Eu.0p=Edo.0p,
              Euo.0p=Euo.0p,
              Lu.0p=Lu.0p,
              file=file))
}

