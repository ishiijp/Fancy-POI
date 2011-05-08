package org.fancypoi


/**
 * Holds the implicit conversions from/to MonadicCondition
 */
object MonadicConversions {

	implicit def bool2Monadic(cond: Boolean) = cond match {
		case true => True
		case _ => False(Nil)
	}

	implicit def monadic2Bool(cond: MonadicCondition): Boolean = cond match {
		case True => true
		case _ => false
	}

}

/**
 * A MonadicCondition allows building boolean expressions of the form
 * (a(0) && a(1) && .. && a(n)), where a(k) is a boolean expression, and
 * collecting the computation failures to a list of messages.
 *
 * <pre>
 * Example:
 *
 *   val isTooYoung = true;
 *   val isTooBad = false;
 *   val isTooStupid = true;
 *
 *   val exp = (!isTooYoung ~ "too young") &&
 *             (!isTooBad ~ "too bad") &&
 *             (!isTooStupid ~ "too stupid")
 *
 *   println(exp match {
 *     case False(msgs) => msgs mkString("Test failed because it is '", "' and '", "'.")
 *     case _ => "success"
 * })
 * </pre>
 */
trait MonadicCondition {
	def &&(cond: MonadicCondition): MonadicCondition

	def ~(msg: String): MonadicCondition
}

case object True extends MonadicCondition {
	def &&(cond: MonadicCondition): MonadicCondition = cond match {
		case f@False(m) => f
		case _ => this
	}

	def ~(msg: String): MonadicCondition = this
}

case class False(msgs: List[String]) extends MonadicCondition {
	def &&(cond: MonadicCondition): MonadicCondition = cond match {
		case False(m) => False(m ::: msgs)
		case _ => this
	}

	def ~(msg: String): MonadicCondition = False(msg :: msgs)
}
