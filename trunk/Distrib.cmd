rmdir \S \Q c:\_temp_dir_
mkdir c:\_temp_dir_
mkdir c:\_temp_dir_\doc
mkdir c:\_temp_dir_\doc\include

copy %1 c:\_temp_dir_\
copy AniSmiley_License.txt c:\_temp_dir_\doc\
copy AniSmiley_changelog.txt c:\_temp_dir_\doc\
copy m_anismiley.h c:\_temp_dir_\doc\include


call "c:\program files\winrar\winrar" A -r -ep1 c:\AniSmiley.zip c:\_temp_dir_\*

rmdir /S /Q c:\_temp_dir_