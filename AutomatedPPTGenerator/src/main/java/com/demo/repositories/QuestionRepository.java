package com.demo.repositories;

import org.springframework.data.jpa.repository.JpaRepository;
import com.demo.entities.Question;

public interface QuestionRepository extends JpaRepository<Question, Long> {
}