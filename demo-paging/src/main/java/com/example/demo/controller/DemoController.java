package com.example.demo.controller;

import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.springframework.data.domain.PageImpl;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.web.PageableDefault;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class DemoController {

	private static final List<String> LIST = IntStream.range(0, 100).mapToObj(i -> String.format("hoge%d", i))
			.collect(Collectors.toList());

	@GetMapping
	public ModelAndView index(@PageableDefault Pageable pageable) {
		final var modelAndView = new ModelAndView("index");
		final var list = LIST.subList(0, 10);
		modelAndView.addObject("list", list);
		final var page = new PageImpl<>(list, pageable, LIST.size());
		modelAndView.addObject("page", page);
		return modelAndView;
	}

	@GetMapping(value = "update")
	public String updateFragment(@RequestParam(value = "keyword", required = false) final String keyword,
			final ModelMap model, @PageableDefault(size = 10) final Pageable pageable) {
		final var filterList = Optional.ofNullable(keyword).isPresent()
				? LIST.stream().filter(str -> str.contains(keyword)).collect(Collectors.toList())
				: LIST;
		final int from = (int) pageable.getOffset();
		final int to = filterList.size() < (pageable.getOffset() + 10) ? filterList.size()
				: (int) (pageable.getOffset() + 10);
		final var list = filterList.subList(from, to);
		model.addAttribute("list", list);
		final var page = new PageImpl<>(list, PageRequest.of(pageable.getPageNumber(), 10, pageable.getSort()),
				filterList.size());
		model.addAttribute("page", page);
		return "list";
	}

	@GetMapping("show")
	public String show() {
		return "img";
	}
	
	@GetMapping("exception")
	public void exception() {
		throw new RuntimeException("exception");
	}
}
