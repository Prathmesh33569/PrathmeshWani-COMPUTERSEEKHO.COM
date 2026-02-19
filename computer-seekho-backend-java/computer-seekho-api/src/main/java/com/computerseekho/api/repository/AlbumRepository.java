//package com.computerseekho.api.repository;
//
//import com.computerseekho.api.entity.Album;
//import com.computerseekho.api.entity.AlbumType;
//import org.springframework.data.jpa.repository.JpaRepository;
//import org.springframework.stereotype.Repository;
//
//import java.util.List;
//@Repository
//public interface AlbumRepository extends JpaRepository<Album, Integer> {
//
//    List<Album> findByAlbumTypeAndAlbumIsActive(
//            AlbumType albumType,
//            Boolean albumIsActive
//    );
//}


package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Album;
import com.computerseekho.api.entity.AlbumType;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;
import com.computerseekho.api.entity.ProgramCode;
import java.util.Optional;
import java.util.List;
@Repository
public interface AlbumRepository extends JpaRepository<Album, Integer> {

    List<Album> findByAlbumTypeAndAlbumIsActive(
            AlbumType albumType,
            Boolean albumIsActive
    );


    Optional<Album> findByAlbumTypeAndProgramCodeIsNull(AlbumType albumType);

    Optional<Album> findByAlbumTypeAndProgramCodeAndAlbumName(
            AlbumType albumType,
            ProgramCode programCode,
            String albumName
    );
}



