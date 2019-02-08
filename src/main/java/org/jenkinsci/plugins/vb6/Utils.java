package org.jenkinsci.plugins.vb6;

import java.io.IOException;
import java.nio.file.Path;
import java.util.stream.Collectors;

public class Utils {
	public static String getContent(Path path) throws IOException{
		String textReturned = null;
		textReturned = java.nio.file.Files.lines(path).collect(Collectors.joining());
		return textReturned;
	}
}
