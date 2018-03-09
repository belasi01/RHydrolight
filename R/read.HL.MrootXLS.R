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
    ndepth <- length(depth)
    waves <- as.numeric(a$wavelen)
    a     <- as.matrix(a[,2:length(x)])

    # Only keep the total absorption coefficient
    ixNA  <- which(is.na(waves))
    nwaves<- ixNA[1] - 1
    waves <- waves[1:nwaves]
    a     <- matrix(as.numeric(a[1:nwaves,]),
                               nrow=nwaves,
                               ncol=ndepth)

    # Scattering
    #####
    b <- read_xlsx(file, sheet = "b", skip=3)
    b <- as.matrix(b[,2:length(x)])
    b     <- matrix(as.numeric(b[1:nwaves,]),
                    nrow=nwaves,
                    ncol=ndepth)

    # Backscattering
    #####
    bb <- read_xlsx(file, sheet = "bb", skip=3)
    bb <- as.matrix(bb[,2:length(x)])
    bb     <- matrix(as.numeric(bb[1:nwaves,]),
                    nrow=nwaves,
                    ncol=ndepth)

    # Backscattering fraction (total)
    #####
    bb.tilde <- read_xlsx(file, sheet = "bb fraction", skip=3)
    bb.tilde <- as.matrix(bb.tilde[,2:length(x)])
    bb.tilde <- bb.tilde[1:nwaves,]
    bb.tilde     <- matrix(as.numeric(bb.tilde[1:nwaves,]),
                     nrow=nwaves,
                     ncol=ndepth)

    # Ed 0+
    #####
    Ed.0p <- read_xlsx(file, sheet = "Ed_in_air", skip=3)
    Ed.0p <- Ed.0p[,-1]
    Ed.0p$Ed_diffuse <- as.numeric(Ed.0p$Ed_diffuse)
    Ed.0p$Ed_total <- as.numeric(Ed.0p$Ed_total)
    Ed.0p$Ed_direct <- as.numeric(Ed.0p$Ed_direct)


    # Lu
    #####
    LuZ <- read_xlsx(file, sheet = "Lu", skip=3)
    x <- colnames(LuZ)
    Lu.0p  <- as.numeric(LuZ$`in air`)
    LuZ <- as.matrix(LuZ[,3:length(x)])
    LuZ <-  matrix(as.numeric(LuZ[1:nwaves,]),
                   nrow=dim(LuZ)[1],
                   ncol=dim(LuZ)[2])

    # EdoZ
    #####
    EdoZ <- read_xlsx(file, sheet = "Eod", skip=3)
    x <- colnames(EdoZ)
    Edo.0p  <- as.numeric(EdoZ$`in air`)
    EdoZ <- as.matrix(EdoZ[,3:length(x)])
    EdoZ <- matrix(as.numeric(EdoZ[1:nwaves,]),
                   nrow=dim(EdoZ)[1],
                   ncol=dim(EdoZ)[2])

    # EuoZ
    #####
    EuoZ <- read_xlsx(file, sheet = "Eou", skip=3)
    x <- colnames(EuoZ)
    Euo.0p  <- as.numeric(EuoZ$`in air`)
    EuoZ <- as.matrix(EuoZ[,3:length(x)])
    EuoZ <- matrix(as.numeric(EuoZ[1:nwaves,]),
                   nrow=dim(EuoZ)[1],
                   ncol=dim(EuoZ)[2])

    # EoZ
    #####
    EoZ <- read_xlsx(file, sheet = "Eo", skip=3)
    x <- colnames(EoZ)
    Eo.0p  <- as.numeric(EoZ$`in air`)
    EoZ <- as.matrix(EoZ[,3:length(x)])
    EoZ <- matrix(as.numeric(EoZ[1:nwaves,]),
                  nrow=dim(EoZ)[1],
                  ncol=dim(EoZ)[2])

    # QoZ
    #####
    QoZ <- read_xlsx(file, sheet = "Eo_quantum", skip=3)
    x <- colnames(QoZ)
    Qo.0p  <- as.numeric(QoZ$`in air`)
    QoZ <- as.matrix(QoZ[,3:length(x)])
    QoZ <- matrix(as.numeric(QoZ[1:nwaves,]),
                  nrow=dim(QoZ)[1],
                  ncol=dim(QoZ)[2])

    # EdZ
    #####
    EdZ <- read_xlsx(file, sheet = "Ed", skip=3)
    x <- colnames(EdZ)
    Ed.0p  <- as.numeric(EdZ$`in air`)
    EdZ <- as.matrix(EdZ[,3:length(x)])
    EdZ <- matrix(as.numeric(EdZ[1:nwaves,]),
                  nrow=dim(EdZ)[1],
                  ncol=dim(EdZ)[2])

    # EuZ
    #####
    EuZ <- read_xlsx(file, sheet = "Eu", skip=3)
    x <- colnames(EuZ)
    Eu.0p  <- as.numeric(EuZ$`in air`)
    EuZ <- as.matrix(EuZ[,3:length(x)])
    EuZ <- matrix(as.numeric(EuZ[1:nwaves,]),
                  nrow=dim(EuZ)[1],
                  ncol=dim(EuZ)[2])

    # R
    R = EuZ/EdZ
    R.0p = Eu.0p/Ed.0p

    x=read_xlsx(file, sheet = "Rrs", skip=3)
    names(x) <- c("waves", "Rrs"    ,  "Ed"      ,   "Lw"     ,   "Lu")
    Rrs.0p <- as.numeric(x$Rrs)
    Lw     <- as.numeric(x$Lw)

    # Kd
    #####
    KdZ <- read_xlsx(file, sheet = "Kd", skip=3)
    x <- colnames(KdZ)
    KdZ <- as.matrix(KdZ[,2:length(x)])
    KdZ <-  matrix(as.numeric(KdZ[1:nwaves,]),
                   nrow=dim(KdZ)[1],
                   ncol=dim(KdZ)[2])

    # Ku
    #####
    KuZ <- read_xlsx(file, sheet = "Ku", skip=3)
    KuZ <- as.matrix(KuZ[,2:length(x)])
    KuZ <-  matrix(as.numeric(KuZ[1:nwaves,]),
                   nrow=dim(KuZ)[1],
                   ncol=dim(KuZ)[2])

    # KLu
    #####
    KLuZ <- read_xlsx(file, sheet = "KLu", skip=3)
    KLuZ <- as.matrix(KLuZ[,2:length(x)])
    KLuZ <-  matrix(as.numeric(KLuZ[1:nwaves,]),
                    nrow=dim(KLuZ)[1],
                    ncol=dim(KLuZ)[2])


    # PAR
    #####
    PAR <- read_xlsx(file, sheet = "PAR", skip=3)
    x <- colnames(PAR)
    PAR <- as.matrix(PAR[,2:length(x)])
    PAR.0p <- as.numeric(PAR[1,])
    PAR.Z <- PAR[-1,]
    PAR.Z <- matrix(as.numeric(PAR.Z), ncol=2)

    # KPAR
    #####
    KPAR <- read_xlsx(file, sheet = "KPAR", skip=3)
    KPAR.Eo <- as.numeric(KPAR$`K_PAR (from Eo)`)

  }


  if (extension == "xls")  {
    # Absorption
    #####
    a <- read_xls(file, sheet = "a", skip=3)
    x <- colnames(a)
    depth <- as.numeric(x[2:length(x)])
    ndepth <- length(depth)
    waves <- as.numeric(a$wavelen)
    a     <- as.matrix(a[,2:length(x)])

    # Only keep the total absorption coefficient
    ixNA  <- which(is.na(waves))
    nwaves<- ixNA[1] - 1
    waves <- waves[1:nwaves]
    a     <- matrix(as.numeric(a[1:nwaves,]),
                    nrow=nwaves,
                    ncol=ndepth)

    # Scattering
    #####
    b <- read_xls(file, sheet = "b", skip=3)
    b <- as.matrix(b[,2:length(x)])
    b     <- matrix(as.numeric(b[1:nwaves,]),
                    nrow=nwaves,
                    ncol=ndepth)

    # Backscattering
    #####
    bb <- read_xls(file, sheet = "bb", skip=3)
    bb <- as.matrix(bb[,2:length(x)])
    bb     <- matrix(as.numeric(bb[1:nwaves,]),
                     nrow=nwaves,
                     ncol=ndepth)

    # Backscattering fraction (total)
    #####
    bb.tilde <- read_xls(file, sheet = "bb fraction", skip=3)
    bb.tilde <- as.matrix(bb.tilde[,2:length(x)])
    bb.tilde <- bb.tilde[1:nwaves,]
    bb.tilde     <- matrix(as.numeric(bb.tilde[1:nwaves,]),
                           nrow=nwaves,
                           ncol=ndepth)

    # Ed 0+
    #####
    Ed.0p <- read_xls(file, sheet = "Ed_in_air", skip=3)
    Ed.0p <- Ed.0p[,-1]
    Ed.0p$Ed_diffuse <- as.numeric(Ed.0p$Ed_diffuse)
    Ed.0p$Ed_total <- as.numeric(Ed.0p$Ed_total)
    Ed.0p$Ed_direct <- as.numeric(Ed.0p$Ed_direct)


    # Lu
    #####
    LuZ <- read_xls(file, sheet = "Lu", skip=3)
    x <- colnames(LuZ)
    Lu.0p  <- as.numeric(LuZ$`in air`)
    LuZ <- as.matrix(LuZ[,3:length(x)])
    LuZ <-  matrix(as.numeric(LuZ[1:nwaves,]),
                   nrow=dim(LuZ)[1],
                   ncol=dim(LuZ)[2])

    # EdoZ
    #####
    EdoZ <- read_xls(file, sheet = "Eod", skip=3)
    x <- colnames(EdoZ)
    Edo.0p  <- as.numeric(EdoZ$`in air`)
    EdoZ <- as.matrix(EdoZ[,3:length(x)])
    EdoZ <- matrix(as.numeric(EdoZ[1:nwaves,]),
                   nrow=dim(EdoZ)[1],
                   ncol=dim(EdoZ)[2])

    # EuoZ
    #####
    EuoZ <- read_xls(file, sheet = "Eou", skip=3)
    x <- colnames(EuoZ)
    Euo.0p  <- as.numeric(EuoZ$`in air`)
    EuoZ <- as.matrix(EuoZ[,3:length(x)])
    EuoZ <- matrix(as.numeric(EuoZ[1:nwaves,]),
                   nrow=dim(EuoZ)[1],
                   ncol=dim(EuoZ)[2])

    # EoZ
    #####
    EoZ <- read_xls(file, sheet = "Eo", skip=3)
    x <- colnames(EoZ)
    Eo.0p  <- as.numeric(EoZ$`in air`)
    EoZ <- as.matrix(EoZ[,3:length(x)])
    EoZ <- matrix(as.numeric(EoZ[1:nwaves,]),
                  nrow=dim(EoZ)[1],
                  ncol=dim(EoZ)[2])

    # QoZ
    #####
    QoZ <- read_xls(file, sheet = "Eo_quantum", skip=3)
    x <- colnames(QoZ)
    Qo.0p  <- as.numeric(QoZ$`in air`)
    QoZ <- as.matrix(QoZ[,3:length(x)])
    QoZ <- matrix(as.numeric(QoZ[1:nwaves,]),
                  nrow=dim(QoZ)[1],
                  ncol=dim(QoZ)[2])

    # EdZ
    #####
    EdZ <- read_xls(file, sheet = "Ed", skip=3)
    x <- colnames(EdZ)
    Ed.0p  <- as.numeric(EdZ$`in air`)
    EdZ <- as.matrix(EdZ[,3:length(x)])
    EdZ <- matrix(as.numeric(EdZ[1:nwaves,]),
                  nrow=dim(EdZ)[1],
                  ncol=dim(EdZ)[2])

    # EuZ
    #####
    EuZ <- read_xls(file, sheet = "Eu", skip=3)
    x <- colnames(EuZ)
    Eu.0p  <- as.numeric(EuZ$`in air`)
    EuZ <- as.matrix(EuZ[,3:length(x)])
    EuZ <- matrix(as.numeric(EuZ[1:nwaves,]),
                  nrow=dim(EuZ)[1],
                  ncol=dim(EuZ)[2])

    # R
    R = EuZ/EdZ
    R.0p = Eu.0p/Ed.0p

    x=read_xls(file, sheet = "Rrs", skip=3)
    names(x) <- c("waves", "Rrs"    ,  "Ed"      ,   "Lw"     ,   "Lu")
    Rrs.0p <- as.numeric(x$Rrs)
    Lw     <- as.numeric(x$Lw)

    # Kd
    #####
    KdZ <- read_xls(file, sheet = "Kd", skip=3)
    x <- colnames(KdZ)
    KdZ <- as.matrix(KdZ[,2:length(x)])
    KdZ <-  matrix(as.numeric(KdZ[1:nwaves,]),
                   nrow=dim(KdZ)[1],
                   ncol=dim(KdZ)[2])

    # Ku
    #####
    KuZ <- read_xls(file, sheet = "Ku", skip=3)
    KuZ <- as.matrix(KuZ[,2:length(x)])
    KuZ <-  matrix(as.numeric(KuZ[1:nwaves,]),
                   nrow=dim(KuZ)[1],
                   ncol=dim(KuZ)[2])

    # KLu
    #####
    KLuZ <- read_xls(file, sheet = "KLu", skip=3)
    KLuZ <- as.matrix(KLuZ[,2:length(x)])
    KLuZ <-  matrix(as.numeric(KLuZ[1:nwaves,]),
                    nrow=dim(KLuZ)[1],
                    ncol=dim(KLuZ)[2])


    # PAR
    #####
    PAR <- read_xls(file, sheet = "PAR", skip=3)
    x <- colnames(PAR)
    PAR <- as.matrix(PAR[,2:length(x)])
    PAR.0p <- as.numeric(PAR[1,])
    PAR.Z <- PAR[-1,]
    PAR.Z <- matrix(as.numeric(PAR.Z), ncol=2)

    # KPAR
    #####
    KPAR <- read_xls(file, sheet = "KPAR", skip=3)
    KPAR.Eo <- as.numeric(KPAR$`K_PAR (from Eo)`)
  }





  return(list(depth=depth, waves=waves,
              a=a,b=b,bb=bb, bb.tilde=bb.tilde,
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

