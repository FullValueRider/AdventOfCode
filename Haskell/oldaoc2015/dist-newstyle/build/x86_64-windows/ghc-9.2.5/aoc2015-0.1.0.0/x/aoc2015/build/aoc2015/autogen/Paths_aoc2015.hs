{-# LANGUAGE CPP #-}
{-# LANGUAGE NoRebindableSyntax #-}
{-# OPTIONS_GHC -fno-warn-missing-import-lists #-}
{-# OPTIONS_GHC -w #-}
module Paths_aoc2015 (
    version,
    getBinDir, getLibDir, getDynLibDir, getDataDir, getLibexecDir,
    getDataFileName, getSysconfDir
  ) where


import qualified Control.Exception as Exception
import qualified Data.List as List
import Data.Version (Version(..))
import System.Environment (getEnv)
import Prelude


#if defined(VERSION_base)

#if MIN_VERSION_base(4,0,0)
catchIO :: IO a -> (Exception.IOException -> IO a) -> IO a
#else
catchIO :: IO a -> (Exception.Exception -> IO a) -> IO a
#endif

#else
catchIO :: IO a -> (Exception.IOException -> IO a) -> IO a
#endif
catchIO = Exception.catch

version :: Version
version = Version [0,1,0,0] []

getDataFileName :: FilePath -> IO FilePath
getDataFileName name = do
  dir <- getDataDir
  return (dir `joinFileName` name)

getBinDir, getLibDir, getDynLibDir, getDataDir, getLibexecDir, getSysconfDir :: IO FilePath



bindir, libdir, dynlibdir, datadir, libexecdir, sysconfdir :: FilePath
bindir     = "c:\\Haskell\\cabal\\bin"
libdir     = "c:\\Haskell\\cabal\\x86_64-windows-ghc-9.2.5\\aoc2015-0.1.0.0-inplace-aoc2015"
dynlibdir  = "c:\\Haskell\\cabal\\x86_64-windows-ghc-9.2.5"
datadir    = "c:\\Haskell\\cabal\\x86_64-windows-ghc-9.2.5\\aoc2015-0.1.0.0"
libexecdir = "c:\\Haskell\\cabal\\aoc2015-0.1.0.0-inplace-aoc2015\\x86_64-windows-ghc-9.2.5\\aoc2015-0.1.0.0"
sysconfdir = "c:\\Haskell\\cabal\\etc"

getBinDir     = catchIO (getEnv "aoc2015_bindir")     (\_ -> return bindir)
getLibDir     = catchIO (getEnv "aoc2015_libdir")     (\_ -> return libdir)
getDynLibDir  = catchIO (getEnv "aoc2015_dynlibdir")  (\_ -> return dynlibdir)
getDataDir    = catchIO (getEnv "aoc2015_datadir")    (\_ -> return datadir)
getLibexecDir = catchIO (getEnv "aoc2015_libexecdir") (\_ -> return libexecdir)
getSysconfDir = catchIO (getEnv "aoc2015_sysconfdir") (\_ -> return sysconfdir)




joinFileName :: String -> String -> FilePath
joinFileName ""  fname = fname
joinFileName "." fname = fname
joinFileName dir ""    = dir
joinFileName dir fname
  | isPathSeparator (List.last dir) = dir ++ fname
  | otherwise                       = dir ++ pathSeparator : fname

pathSeparator :: Char
pathSeparator = '\\'

isPathSeparator :: Char -> Bool
isPathSeparator c = c == '/' || c == '\\'
