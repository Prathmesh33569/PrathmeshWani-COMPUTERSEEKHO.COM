package com.computerseekho.api.repository;

import com.computerseekho.api.entity.Image;
import org.springframework.data.jpa.repository.*;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface RecruitersRepository extends JpaRepository<Image, Integer> {

    @Query(value = """
        SELECT im.image_path
        FROM image_master im
        WHERE im.image_is_active = 1
          AND im.album_id = (
              SELECT am.album_id
              FROM album_master am
              WHERE am.album_name = 'Recruiters'
                AND am.album_is_active = 1
          )
        """, nativeQuery = true)
    List<String> findRecruiterImages();
}
